/**
 * Enhanced DocxHandler - Renders DOCX files to Canvas with proper pagination
 * This implementation preserves both explicit and implicit page breaks
 */
function DocxHandler(canvas) {
    var self = this;
    self.DPI = 96;

    var docxFile;
    var currPageIndex = 0;
    var documentPages = []; // Will store rendered pages
    var pageMetrics = []; // Will store page size and break information
    var documentLoaded = false;
    
    // Settings for pagination calculation
    var pageSettings = {
        width: 8.5 * self.DPI, // Default letter size in pixels
        height: 11 * self.DPI,
        margins: {
            top: self.DPI, // 1 inch margins
            right: self.DPI,
            bottom: self.DPI,
            left: self.DPI
        }
    };

    /**
     * Load the DOCX document from a URL
     */
    this.loadDocument = function(documentUrl, onCompletion, onError) {
        benchmark.time("Document loaded");
        var xhr = new XMLHttpRequest();
        xhr.open('GET', documentUrl, true);
        xhr.responseType = 'arraybuffer';

        xhr.onload = function () {
            if (xhr.status === 200) {
                docxFile = xhr.response;
                benchmark.timeEnd("Document loaded");
                self.processDocument().then(() => {
                    if (onCompletion) onCompletion();
                }).catch(error => {
                    console.error("Error processing document:", error);
                    if (onError) onError(error);
                });
            } else {
                const msg = "Failed to load document. Status: " + xhr.status;
                console.error(msg);
                if (onError) onError(msg);
            }
        };

        xhr.onerror = function () {
            const msg = "Network error while loading document";
            console.error(msg);
            if (onError) onError(msg);
        };

        xhr.send();
    };

    /**
     * Process the document and extract pages with pagination
     */
    this.processDocument = function() {
        return new Promise((resolve, reject) => {
            try {
                // Use docx-preview directly for pagination
                // This approach uses the native docx-preview library which already supports pagination
                benchmark.time("Convert to paginated HTML");
                
                // Create container for the docx-preview output
                const tempContainer = document.createElement('div');
                tempContainer.style.position = 'absolute';
                tempContainer.style.left = '-9999px';
                document.body.appendChild(tempContainer);
                
                // Use docx-preview to render the document
                docx.renderAsync(docxFile, tempContainer, null, {
                    className: 'docx',
                    inWrapper: true,
                    ignoreWidth: false,
                    ignoreHeight: false,
                    ignoreFonts: false,
                    breakPages: true, // Enable pagination
                    renderHeaders: true,
                    renderFooters: true,
                    renderFootnotes: true,
                    renderEndnotes: true,
                    useBase64URL: true
                }).then(result => {
                    benchmark.timeEnd("Convert to paginated HTML");
                    
                    // Find all page divs
                    const pageElements = tempContainer.querySelectorAll('.docx-page');
                    console.log(`Found ${pageElements.length} pages in the document`);
                    
                    if (pageElements.length === 0) {
                        // Fallback to mammoth if docx-preview doesn't paginate
                        processMammothFallback().then(resolve).catch(reject);
                        document.body.removeChild(tempContainer);
                        return;
                    }
                    
                    // Extract page dimensions from the first page
                    if (pageElements.length > 0) {
                        const firstPage = pageElements[0];
                        const style = window.getComputedStyle(firstPage);
                        pageSettings.width = parseFloat(style.width);
                        pageSettings.height = parseFloat(style.height);
                        console.log(`Page dimensions: ${pageSettings.width}x${pageSettings.height}`);
                    }
                    
                    // Process each page
                    benchmark.time("Render pages to canvas");
                    let processedPages = 0;
                    
                    // Function to render a page element to canvas
                    const renderPageElement = (pageElement, index) => {
                        return new Promise((resolve, reject) => {
                            // Use html2canvas to render the page
                            html2canvas(pageElement, {
                                width: pageSettings.width,
                                height: pageSettings.height,
                                useCORS: true,
                                allowTaint: true,
                                backgroundColor: '#ffffff',
                                scale: 1,
                                logging: false,
                                onclone: function(clonedDoc) {
                                    // Ensure images are properly loaded in the clone
                                    const images = clonedDoc.querySelectorAll('img');
                                    Array.from(images).forEach(img => {
                                        img.crossOrigin = 'anonymous';
                                    });
                                }
                            }).then(pageCanvas => {
                                documentPages[index] = pageCanvas;
                                resolve();
                            }).catch(error => {
                                console.error(`Error rendering page ${index}:`, error);
                                // Create a blank page as fallback
                                const blankCanvas = document.createElement('canvas');
                                blankCanvas.width = pageSettings.width;
                                blankCanvas.height = pageSettings.height;
                                const ctx = blankCanvas.getContext('2d');
                                ctx.fillStyle = '#ffffff';
                                ctx.fillRect(0, 0, blankCanvas.width, blankCanvas.height);
                                ctx.fillStyle = '#000000';
                                ctx.font = '14px Arial';
                                ctx.fillText(`Error rendering page ${index + 1}`, 20, 20);
                                documentPages[index] = blankCanvas;
                                resolve();
                            });
                        });
                    };
                    
                    // Process pages in batches to avoid memory issues
                    const batchSize = 3;
                    const totalPages = pageElements.length;
                    let currentBatch = 0;
                    
                    const processNextBatch = () => {
                        const startIdx = currentBatch * batchSize;
                        const endIdx = Math.min(startIdx + batchSize, totalPages);
                        
                        if (startIdx >= totalPages) {
                            // All batches processed
                            benchmark.timeEnd("Render pages to canvas");
                            document.body.removeChild(tempContainer);
                            documentLoaded = true;
                            resolve();
                            return;
                        }
                        
                        console.log(`Processing batch ${currentBatch + 1}, pages ${startIdx + 1}-${endIdx}`);
                        
                        // Create promises for current batch
                        const batchPromises = [];
                        for (let i = startIdx; i < endIdx; i++) {
                            batchPromises.push(renderPageElement(pageElements[i], i));
                        }
                        
                        // Process current batch then move to next
                        Promise.all(batchPromises).then(() => {
                            currentBatch++;
                            setTimeout(processNextBatch, 0); // Use setTimeout to prevent stack overflow
                        }).catch(error => {
                            console.error("Error processing batch:", error);
                            currentBatch++;
                            setTimeout(processNextBatch, 0);
                        });
                    };
                    
                    // Start batch processing
                    processNextBatch();
                    
                }).catch(error => {
                    console.error("Error with docx-preview, falling back to mammoth:", error);
                    document.body.removeChild(tempContainer);
                    processMammothFallback().then(resolve).catch(reject);
                });
                
            } catch (error) {
                console.error("Error in processDocument:", error);
                processMammothFallback().then(resolve).catch(reject);
            }
        });
    };
    
    /**
     * Fallback method using mammoth.js
     */
    function processMammothFallback() {
        return new Promise((resolve, reject) => {
            console.log("Using mammoth.js fallback for document processing");
            benchmark.time("Mammoth conversion");
            
            mammoth.convertToHtml({arrayBuffer: docxFile}, {
                includeDefaultStyleMap: true,
                styleMap: [
                    "w:br[type='page'] => hr.page-break",
                    "p[style-name='Heading 1'] => h1:fresh",
                    "p[style-name='Heading 2'] => h2:fresh",
                    "table => table.docx-table"
                ]
            }).then(function(result) {
                benchmark.timeEnd("Mammoth conversion");
                
                if (result.messages.length > 0) {
                    console.warn("Mammoth warnings:", result.messages);
                }
                
                // Create a container for the HTML content
                const container = document.createElement('div');
                container.style.position = 'absolute';
                container.style.left = '-9999px';
                container.style.width = `${pageSettings.width - pageSettings.margins.left - pageSettings.margins.right}px`;
                container.style.fontFamily = 'Arial, sans-serif';
                container.style.fontSize = '12pt';
                container.innerHTML = result.value;
                document.body.appendChild(container);
                
                // Find all explicit page breaks in the HTML
                const explicitBreakElements = container.querySelectorAll('hr.page-break');
                console.log(`Found ${explicitBreakElements.length} explicit page breaks in the document`);
                
                // Calculate content height for implicit pagination
                const contentHeight = container.scrollHeight;
                const contentPerPage = pageSettings.height - pageSettings.margins.top - pageSettings.margins.bottom;
                const estimatedPageCount = Math.max(1, Math.ceil(contentHeight / contentPerPage));
                
                console.log(`Content height: ${contentHeight}px, available height per page: ${contentPerPage}px`);
                console.log(`Estimated page count: ${estimatedPageCount}`);
                
                // Create pages based on content height if no explicit breaks
                if (explicitBreakElements.length === 0) {
                    benchmark.time("Render pages");
                    
                    // Create page canvases
                    const pages = [];
                    let processedPages = 0;
                    
                    for (let i = 0; i < estimatedPageCount; i++) {
                        // Clone container for each page
                        const pageDiv = container.cloneNode(true);
                        pageDiv.style.height = `${contentPerPage}px`;
                        pageDiv.style.overflow = 'hidden';
                        pageDiv.style.position = 'relative';
                        pageDiv.style.top = `-${i * contentPerPage}px`;
                        pageDiv.style.clip = `rect(${i * contentPerPage}px, ${pageSettings.width}px, ${(i + 1) * contentPerPage}px, 0)`;
                        document.body.appendChild(pageDiv);
                        
                        // Render to canvas
                        html2canvas(pageDiv, {
                            width: pageSettings.width - pageSettings.margins.left - pageSettings.margins.right,
                            height: contentPerPage,
                            useCORS: true,
                            allowTaint: true,
                            backgroundColor: '#ffffff'
                        }).then(function(renderedCanvas) {
                            // Create final page canvas with margins
                            const pageCanvas = document.createElement('canvas');
                            pageCanvas.width = pageSettings.width;
                            pageCanvas.height = pageSettings.height;
                            
                            const ctx = pageCanvas.getContext('2d');
                            ctx.fillStyle = '#ffffff';
                            ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
                            
                            // Draw content with margins
                            ctx.drawImage(
                                renderedCanvas,
                                pageSettings.margins.left,
                                pageSettings.margins.top
                            );
                            
                            pages[i] = pageCanvas;
                            document.body.removeChild(pageDiv);
                            
                            processedPages++;
                            if (processedPages === estimatedPageCount) {
                                documentPages = pages;
                                document.body.removeChild(container);
                                benchmark.timeEnd("Render pages");
                                documentLoaded = true;
                                resolve();
                            }
                        }).catch(function(error) {
                            console.error(`Error rendering page ${i}:`, error);
                            document.body.removeChild(pageDiv);
                            
                            // Create blank page as fallback
                            const blankCanvas = document.createElement('canvas');
                            blankCanvas.width = pageSettings.width;
                            blankCanvas.height = pageSettings.height;
                            const ctx = blankCanvas.getContext('2d');
                            ctx.fillStyle = '#ffffff';
                            ctx.fillRect(0, 0, blankCanvas.width, blankCanvas.height);
                            
                            pages[i] = blankCanvas;
                            
                            processedPages++;
                            if (processedPages === estimatedPageCount) {
                                documentPages = pages;
                                document.body.removeChild(container);
                                benchmark.timeEnd("Render pages");
                                documentLoaded = true;
                                resolve();
                            }
                        });
                    }
                } else {
                    // Handle explicit page breaks
                    benchmark.time("Process explicit breaks");
                    
                    // Split document at page breaks
                    const pageContents = [];
                    let currentContent = document.createElement('div');
                    
                    // Process all child nodes
                    Array.from(container.childNodes).forEach(node => {
                        // If node is a page break, start a new page
                        if (node.nodeName === 'HR' && node.classList.contains('page-break')) {
                            pageContents.push(currentContent);
                            currentContent = document.createElement('div');
                        } else {
                            // Clone the node and add to current content
                            currentContent.appendChild(node.cloneNode(true));
                        }
                    });
                    
                    // Add the last page
                    pageContents.push(currentContent);
                    
                    console.log(`Split content into ${pageContents.length} pages`);
                    
                    // Render each page
                    const pages = [];
                    let processedPages = 0;
                    
                    pageContents.forEach((content, index) => {
                        // Style the content div
                        content.style.position = 'absolute';
                        content.style.left = '-9999px';
                        content.style.width = `${pageSettings.width - pageSettings.margins.left - pageSettings.margins.right}px`;
                        content.style.fontFamily = 'Arial, sans-serif';
                        content.style.fontSize = '12pt';
                        document.body.appendChild(content);
                        
                        // Render to canvas
                        html2canvas(content, {
                            width: pageSettings.width - pageSettings.margins.left - pageSettings.margins.right,
                            height: content.scrollHeight,
                            useCORS: true,
                            allowTaint: true,
                            backgroundColor: '#ffffff'
                        }).then(function(renderedCanvas) {
                            // Create final page canvas with margins
                            const pageCanvas = document.createElement('canvas');
                            pageCanvas.width = pageSettings.width;
                            pageCanvas.height = pageSettings.height;
                            
                            const ctx = pageCanvas.getContext('2d');
                            ctx.fillStyle = '#ffffff';
                            ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
                            
                            // Draw content with margins (and scale if content is too tall)
                            const contentHeight = renderedCanvas.height;
                            const availableHeight = pageSettings.height - pageSettings.margins.top - pageSettings.margins.bottom;
                            
                            if (contentHeight > availableHeight) {
                                // Scale down to fit
                                const scale = availableHeight / contentHeight;
                                const scaledWidth = renderedCanvas.width * scale;
                                
                                ctx.drawImage(
                                    renderedCanvas,
                                    pageSettings.margins.left,
                                    pageSettings.margins.top,
                                    scaledWidth,
                                    availableHeight
                                );
                            } else {
                                // No scaling needed
                                ctx.drawImage(
                                    renderedCanvas,
                                    pageSettings.margins.left,
                                    pageSettings.margins.top
                                );
                            }
                            
                            pages[index] = pageCanvas;
                            document.body.removeChild(content);
                            
                            processedPages++;
                            if (processedPages === pageContents.length) {
                                documentPages = pages;
                                document.body.removeChild(container);
                                benchmark.timeEnd("Process explicit breaks");
                                documentLoaded = true;
                                resolve();
                            }
                        }).catch(function(error) {
                            console.error(`Error rendering page ${index}:`, error);
                            document.body.removeChild(content);
                            
                            // Create blank page as fallback
                            const blankCanvas = document.createElement('canvas');
                            blankCanvas.width = pageSettings.width;
                            blankCanvas.height = pageSettings.height;
                            const ctx = blankCanvas.getContext('2d');
                            ctx.fillStyle = '#ffffff';
                            ctx.fillRect(0, 0, blankCanvas.width, blankCanvas.height);
                            
                            pages[index] = blankCanvas;
                            
                            processedPages++;
                            if (processedPages === pageContents.length) {
                                documentPages = pages;
                                document.body.removeChild(container);
                                benchmark.timeEnd("Process explicit breaks");
                                documentLoaded = true;
                                resolve();
                            }
                        });
                    });
                }
            }).catch(function(error) {
                console.error("Error in mammoth conversion:", error);
                reject(error);
            });
        });
    }

    /**
     * Return the number of pages in the document
     */
    this.pageCount = function() {
        return documentPages.length || 1;
    };

    /**
     * Return the original width of the document
     */
    this.originalWidth = function() {
        return pageSettings.width;
    };

    /**
     * Return the original height of the document
     */
    this.originalHeight = function() {
        return pageSettings.height;
    };

    /**
     * Draw the document to the canvas
     */
    this.drawDocument = function(scale, rotation, onCompletion) {
        benchmark.time("Document drawn");
        self.redraw(scale, rotation, currPageIndex, function(err) {
            benchmark.timeEnd("Document drawn");
            if (onCompletion) onCompletion(err);
        });
    };

    /**
     * Redraw the document with the specified parameters
     */
    this.redraw = function(scale, rotation, index, onCompletion) {
        if (!documentLoaded || documentPages.length === 0) {
            console.error("Document not loaded or has no pages");
            if (onCompletion) onCompletion(new Error("Document not loaded or has no pages"));
            return;
        }

        // Ensure valid index
        index = Math.max(0, Math.min(index, documentPages.length - 1));
        currPageIndex = index;
        
        // Get the page canvas for the current index
        const pageCanvas = documentPages[index];
        
        if (!pageCanvas) {
            console.error(`Page ${index} not rendered`);
            if (onCompletion) onCompletion(new Error(`Page ${index} not rendered`));
            return;
        }
        
        // Draw the page to the main canvas with scaling and rotation
        const ctx = canvas.getContext('2d');
        
        // Set canvas dimensions
        const canvasWidth = pageCanvas.width * scale;
        const canvasHeight = pageCanvas.height * scale;
        
        canvas.width = canvasWidth;
        canvas.height = canvasHeight;
        
        // Clear the canvas
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        // Apply rotation if needed
        if (rotation !== 0) {
            ctx.save();
            ctx.translate(canvas.width / 2, canvas.height / 2);
            ctx.rotate(rotation * Math.PI / 180);
            ctx.scale(scale, scale);
            ctx.drawImage(pageCanvas, -pageCanvas.width / 2, -pageCanvas.height / 2);
            ctx.restore();
        } else {
            ctx.scale(scale, scale);
            ctx.drawImage(pageCanvas, 0, 0);
        }
        
        if (onCompletion) onCompletion();
    };

    /**
     * Apply a function to the canvas
     */
    this.applyToCanvas = function(apply) {
        apply(canvas);
    };

    /**
     * Create canvases for multiple pages
     */
    this.createCanvases = function(callback, dpiCalcFunction, fromPage, pageCount) {
        const canvases = [];
        const toPage = Math.min(documentPages.length, fromPage + pageCount - 1);
        let completed = 0;

        for (let i = fromPage; i <= toPage; i++) {
            this.redraw(1, 0, i, function() {
                canvases.push({
                    canvas: canvas.cloneNode(true),
                    originalDocumentDpi: self.DPI
                });
                if (++completed === (toPage - fromPage + 1)) {
                    callback(canvases);
                }
            });
        }
    };
}
