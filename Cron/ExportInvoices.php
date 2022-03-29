<?php

declare(strict_types=1);

namespace DevAll\ExportInvoices\Cron;

use Magento\Framework\App\Area;
use Magento\Framework\App\Filesystem\DirectoryList;
use Magento\Framework\Exception\FileSystemException;
use Magento\Framework\Filesystem;
use Magento\Framework\Message\ManagerInterface;
use Magento\Framework\Stdlib\DateTime\DateTime;
use Magento\Sales\Api\OrderAddressRepositoryInterface;
use Magento\Sales\Model\ResourceModel\Order\Invoice\CollectionFactory;
use Magento\Sales\Model\ResourceModel\Order\Creditmemo\CollectionFactory as CreditCollectionFactory;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use Snmportal\Pdfprint\Helper\Template;
use Magento\Framework\App\Config\ScopeConfigInterface;
use Psr\Log\LoggerInterface;
use Zend\Mime\Message;
use Zend\Mime\Part;
use ZipArchive;
use Magento\Framework\Filesystem\DriverInterface;
use Magento\Framework\Mail\TransportInterfaceFactory;
use Magento\Store\Model\StoreManagerInterface;
use Magento\Framework\Translate\Inline\StateInterface;
use DevAll\ExportInvoices\Model\Mail\Template\TransportBuilder;
use Zend\Mime\Mime;

class ExportInvoices
{
    /**
     * @var CollectionFactory
     */
    private $collection;

    /**
     * @var DateTime
     */
    private $date;

    /**
     * @var DateTime
     */
    private $stdlibDateTime;

    /**
     * @var OrderAddressRepositoryInterface
     */
    private $orderAddressRepository;

    /**
     * @var Filesystem
     */
    private $filesystem;

    /**
     * @var Template
     */
    private $templateHelper;

    /**
     * @var ScopeConfigInterface
     */
    private $scopeConfig;

    /**
     * @var LoggerInterface
     */
    private $logger;

    /**
     * @var DriverInterface
     */
    private $driver;

    /**
     * @var string
     */
    private $pfad;

    /**
     * @var TransportInterfaceFactory
     */
    private $mailTransportFactory;

    /**
     * @var StoreManagerInterface
     */
    private $storeManager;

    /**
     * @var DriverInterface
     */
    private $driverInterface;

    /**
     * @var CreditCollectionFactory
     */
    private $creditCollection;

    /**
     * @var ManagerInterface
     */
    private $_messageManager;

    /**
     * @var TransportBuilder
     */
    private $_transportBuilder;

    /**
     * @var StateInterface
     */
    private $inlineTranslation;

    /**
     * @param StateInterface $inlineTranslation
     * @param TransportBuilder $transportBuilder
     * @param ManagerInterface $messageManager
     * @param TransportInterfaceFactory $mailTransportFactory
     * @param CollectionFactory $collection
     * @param CreditCollectionFactory $creditCollection
     * @param DateTime $stdlibDateTime
     * @param DateTime $date
     * @param OrderAddressRepositoryInterface $orderAddressRepository
     * @param Template $templateHelper
     * @param Filesystem $filesystem
     * @param ScopeConfigInterface $scopeConfig
     * @param DriverInterface $driver
     * @param LoggerInterface $logger
     * @param StoreManagerInterface $storeManager
     * @param DriverInterface $driverInterface
     */
    public function __construct(
        StateInterface $inlineTranslation,
        TransportBuilder $transportBuilder,
        ManagerInterface $messageManager,
        TransportInterfaceFactory $mailTransportFactory,
        CollectionFactory $collection,
        CreditCollectionFactory $creditCollection,
        DateTime $stdlibDateTime,
        DateTime $date,
        OrderAddressRepositoryInterface $orderAddressRepository,
        Template $templateHelper,
        Filesystem $filesystem,
        ScopeConfigInterface $scopeConfig,
        DriverInterface $driver,
        LoggerInterface $logger,
        StoreManagerInterface $storeManager,
        DriverInterface $driverInterface
    ) {
        $this->inlineTranslation = $inlineTranslation;
        $this->_transportBuilder = $transportBuilder;
        $this->_messageManager = $messageManager;
        $this->collection = $collection;
        $this->creditCollection = $creditCollection;
        $this->stdlibDateTime = $stdlibDateTime;
        $this->date = $date;
        $this->orderAddressRepository = $orderAddressRepository;
        $this->templateHelper = $templateHelper;
        $this->filesystem = $filesystem;
        $this->scopeConfig = $scopeConfig;
        $this->logger = $logger;
        $this->driver = $driver;
        $this->mailTransportFactory = $mailTransportFactory;
        $this->storeManager = $storeManager;
        $this->driverInterface = $driverInterface;
    }

    /**
     * Export non german invoices
     */
    public function execute(): void
    {
        $this->pfad = '';
        $currentTime = $this->stdlibDateTime->date();
        $lastMonth = $this->date->gmtDate('Y-m-d h:m:s\Z', strtotime('-1 month'));
        $invoiceCollection = $this->collection->create()
            ->addAttributeToFilter('created_at', ['from'=>$lastMonth, 'to'=>$currentTime])
            ->getItems();
        $creditMemoCollection = $this->creditCollection->create()
            ->addAttributeToFilter('created_at', ['from'=>$lastMonth, 'to'=>$currentTime])
            ->getItems();

        foreach ($invoiceCollection as $invoice) {
            $countryCode = $this->orderAddressRepository->get($invoice->getShippingAddressId())->getCountryId();
            if ($countryCode !== 'DE') {
                $this->generateAndSaveInvoicePdf($invoice);
            }
        }

        //add pdfs to zip archive
        $saveDir = $this->pfad;
        $rootPath = $this->driverInterface->getRealPath($saveDir);
        if ($rootPath) {

            // Initialize archive object
            $zip = new ZipArchive();
            $zip->open($rootPath . '.zip', ZipArchive::CREATE | ZipArchive::OVERWRITE);
            // Create recursive directory iterator
            /** @var SplFileInfo[] $files */
            $files = new RecursiveIteratorIterator(
                new RecursiveDirectoryIterator($rootPath),
                RecursiveIteratorIterator::LEAVES_ONLY
            );

            foreach ($files as $name => $file) {
                // Skip directories (they would be added automatically)
                if (!$file->isDir()) {
                    // Get real and relative path for current file
                    $filePath = $file->getRealPath();
                    $relativePath = substr($filePath, strlen($rootPath) + 1);

                    // Add current file to archive
                    $zip->addFile($filePath, $relativePath);
                }
            }
            $zip->close();

            //deletes directory after creating archive
            $this->driverInterface->deleteDirectory($rootPath);
        }

        //send zip download link via email
        $message = new \Magento\Framework\Mail\Message();
        $message->setFrom('mail.recipient@example.com');
        $message->addTo($this->scopeConfig->getValue('invoice_export_pdf/invoice/send_to'));
        $message->setSubject('Monthly PDF Invoices Download Link (Non-German)');
        $message->setBody($this->storeManager->getStore()->getBaseUrl() . substr($this->pfad, 4)  . '.zip');
        $transport = $this->mailTransportFactory->create(['message' => $message]);
        $transport->sendMessage();

        foreach ($creditMemoCollection as $creditMemo) {
            $this->generateAndSaveInvoicePdf($creditMemo);
            $creditMemoId = $creditMemo->getIncrementId();

            $sendTo = $this->scopeConfig->getValue('invoice_export_pdf/invoice/send_to');

            $this->inlineTranslation->suspend();

            $transport =
                $this->_transportBuilder
                    ->setTemplateIdentifier('devall_email_with_pdf')
                    ->setTemplateOptions(['area' => Area::AREA_FRONTEND, 'store' => 1])
                    ->setTemplateVars([
                            'fileId' => $creditMemoId,
                            'fileType' => 'credit memo'
                        ])
                    ->setFrom('general')
                    ->addTo($sendTo)
                    ->getTransport();

            $html = $transport->getMessage()->getBody()->generateMessage('');
            $bodyMessage = new Part(Mime::encode(quoted_printable_decode($html), 'utf-8', "\n"));
            $bodyMessage->type = 'text/html';
            $pdfFile = $this->driverInterface->fileGetContents($saveDir . '/' . $creditMemoId . '.pdf');
            $attachment = $this->_transportBuilder->addAttachment($pdfFile, $creditMemoId . '.pdf');
            $bodyPart = new Message();
            $bodyPart->setParts([$bodyMessage,$attachment]);
            $transport->getMessage()->setBody($bodyPart);
            $transport->sendMessage();
            $this->inlineTranslation->resume();
        }

        foreach ($invoiceCollection as $invoice) {
            $this->generateAndSaveInvoicePdf($invoice);
            $invoiceId = $invoice->getIncrementId();
            $sendTo = $this->scopeConfig->getValue('invoice_export_pdf/invoice/send_to');
            $this->inlineTranslation->suspend();
            $transport =
                $this->_transportBuilder
                    ->setTemplateIdentifier('devall_email_with_pdf')
                    ->setTemplateOptions(['area' => Area::AREA_FRONTEND, 'store' => 1])
                    ->setTemplateVars([
                            'fileId' => $invoiceId,
                            'fileType' => 'invoice'
                        ])
                    ->setFrom('general')
                    ->addTo($sendTo)
                    ->getTransport();

            $html = $transport->getMessage()->getBody()->generateMessage('');
            $bodyMessage = new Part(Mime::encode(quoted_printable_decode($html), 'utf-8', "\n"));
            $bodyMessage->type = 'text/html';
            $pdfFile = $this->driverInterface->fileGetContents($saveDir . '/' . $invoiceId . '.pdf');
            $attachment = $this->_transportBuilder->addAttachment($pdfFile, $invoiceId . '.pdf');
            $bodyPart = new Message();
            $bodyPart->setParts([$bodyMessage,$attachment]);
            $transport->getMessage()->setBody($bodyPart);
            $transport->sendMessage();
            $this->inlineTranslation->resume();
        }
    }

    /**
     * Generate Invoice or Credit memo pdf
     *
     * @param object $invoice
     * @return void
     */
    public function generateAndSaveInvoicePdf($invoice): void
    {
        $currentMonth = $this->stdlibDateTime->date('Y-m');
        $saveFile = 'pub/media/pdf-documents/{{var entity. increment_id}}';
        if ($saveFile) {
            try {
                $baseDirectory = $this->filesystem->getDirectoryWrite(DirectoryList::ROOT);
                $filename = $saveFile . '.pdf';
                $engine = $this->templateHelper->getEngineForDocument($invoice);
                if ($engine) {
                    $pdf = $engine->getPdf([$invoice]);
                    $filename = $engine->filterValue($filename);
                    $pfad = $this->driver->getParentDirectory($filename).'/'.$currentMonth;
                    $this->pfad = $pfad;
                    if (!$baseDirectory->isWritable($pfad)) {
                        $baseDirectory->create($pfad);
                    }
                    $baseDirectory->writeFile($pfad.'/'.$invoice->getIncrementId().'.pdf', $pdf->render());
                }
            } catch (FileSystemException $e) {
                $this->logger->log('Missing archive invoice archive directory path', $e->getLogMessage());
            }
        }
    }
}
