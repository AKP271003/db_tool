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
                // First, use mammoth.js to extract the document structure
                benchmark.time("Extract DOCX content");
                
                // Create a zip to extract OOXML content directly
                JSZip.loadAsync(docxFile).then(function(zip) {
                    // Get document.xml content (main content)
                    zip.file("word/document.xml").async("string").then(function(documentXml) {
                        // Get styles.xml (for paragraph styles)
                        zip.file("word/styles.xml").async("string").then(function(stylesXml) {
                            // Get settings.xml (for document settings)
                            zip.file("word/settings.xml").async("string").then(function(settingsXml) {
                                // Process document structure to identify page breaks
                                processDocumentStructure(documentXml, stylesXml, settingsXml).then(() => {
                                    benchmark.timeEnd("Extract DOCX content");
                                    
                                    // Now convert to HTML for rendering
                                    benchmark.time("Convert to HTML");
                                    mammoth.convertToHtml({arrayBuffer: docxFile}, {
                                        includeEmbeddedStyleMap: true,
                                        includeDefaultStyleMap: true,
                                        styleMap: [
                                            "w:br[type='page'] => hr.page-break",
                                            "p[style-name='Heading 1'] => h1:fresh",
                                            "p[style-name='Heading 2'] => h2:fresh",
                                            "table => table.docx-table"
                                        ]
                                    }).then(function(result) {
                                        benchmark.timeEnd("Convert to HTML");
                                        
                                        // Process warnings if any
                                        if (result.messages.length > 0) {
                                            console.warn("Mammoth warnings:", result.messages);
                                        }
                                        
                                        // Split HTML content by page breaks and render each page
                                        benchmark.time("Split and render pages");
                                        renderPaginatedContent(result.value).then(() => {
                                            benchmark.timeEnd("Split and render pages");
                                            documentLoaded = true;
                                            resolve();
                                        }).catch(reject);
                                    }).catch(reject);
                                }).catch(reject);
                            }).catch(reject);
                        }).catch(reject);
                    }).catch(reject);
                }).catch(reject);
            } catch (error) {
                reject(error);
            }
        });
    };

    /**
     * Process the DOCX XML structure to identify explicit and implicit page breaks
     */
    function processDocumentStructure(documentXml, stylesXml, settingsXml) {
        return new Promise((resolve) => {
            // Parse XML documents
            const parser = new DOMParser();
            const docXml = parser.parseFromString(documentXml, "text/xml");
            const stylesDoc = parser.parseFromString(stylesXml, "text/xml");
            
            // Extract page size and margins from section properties
            extractPageSettings(docXml);
            
            // Find explicit page breaks
            const pageBreaks = findExplicitPageBreaks(docXml);
            
            // Find section breaks (also cause page breaks)
            const sectionBreaks = findSectionBreaks(docXml);
            
            // Combine explicit page breaks and section breaks
            const allBreaks = [...pageBreaks, ...sectionBreaks].sort((a, b) => a.position - b.position);
            
            console.log(`Found ${pageBreaks.length} explicit page breaks and ${sectionBreaks.length} section breaks`);
            
            // Store breaks for pagination
            pageMetrics = {
                pageSize: pageSettings,
                breaks: allBreaks
            };
            
            resolve();
        });
    }
    
    /**
     * Extract page size and margins from the document
     */
    function extractPageSettings(docXml) {
        // Find all section properties
        const sectPrs = docXml.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "sectPr");
        
        if (sectPrs.length > 0) {
            // Use the first section's properties (TODO: handle multiple sections with different page sizes)
            const sectPr = sectPrs[0];
            
            // Get page size
            const pgSz = sectPr.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "pgSz")[0];
            if (pgSz) {
                // Width and height are in twentieths of a point
                const widthTwips = parseInt(pgSz.getAttribute("w:w") || "12240"); // Default to 8.5"
                const heightTwips = parseInt(pgSz.getAttribute("w:h") || "15840"); // Default to 11"
                
                // Convert to pixels (1 twip = 1/1440 inch, 1 inch = 96 pixels at 96 DPI)
                pageSettings.width = (widthTwips * self.DPI) / 1440;
                pageSettings.height = (heightTwips * self.DPI) / 1440;
            }
            
            // Get page margins
            const pgMar = sectPr.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "pgMar")[0];
            if (pgMar) {
                // Margins are in twentieths of a point
                const topTwips = parseInt(pgMar.getAttribute("w:top") || "1440"); // Default to 1"
                const rightTwips = parseInt(pgMar.getAttribute("w:right") || "1440");
                const bottomTwips = parseInt(pgMar.getAttribute("w:bottom") || "1440");
                const leftTwips = parseInt(pgMar.getAttribute("w:left") || "1440");
                
                // Convert to pixels
                pageSettings.margins.top = (topTwips * self.DPI) / 1440;
                pageSettings.margins.right = (rightTwips * self.DPI) / 1440;
                pageSettings.margins.bottom = (bottomTwips * self.DPI) / 1440;
                pageSettings.margins.left = (leftTwips * self.DPI) / 1440;
            }
        }
        
        console.log("Page settings:", pageSettings);
    }
    
    /**
     * Find explicit page breaks in the document
     */
    function findExplicitPageBreaks(docXml) {
        const breaks = [];
        const brElements = docXml.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "br");
        
        let position = 0;
        for (let i = 0; i < brElements.length; i++) {
            const br = brElements[i];
            if (br.getAttribute("w:type") === "page") {
                // Find the parent paragraph for position tracking
                let parent = br.parentNode;
                while (parent && parent.nodeName !== "w:p") {
                    parent = parent.parentNode;
                }
                
                // Calculate position based on preceding elements
                if (parent) {
                    const paragraphs = docXml.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p");
                    for (let j = 0; j < paragraphs.length; j++) {
                        if (paragraphs[j] === parent) {
                            position = j;
                            break;
                        }
                    }
                }
                
                breaks.push({
                    type: "explicit",
                    position: position,
                    element: br
                });
            }
        }
        
        return breaks;
    }
    
    /**
     * Find section breaks in the document
     */
    function findSectionBreaks(docXml) {
        const breaks = [];
        const sectPrs = docXml.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "sectPr");
        
        for (let i = 0; i < sectPrs.length; i++) {
            const sectPr = sectPrs[i];
            
            // Find parent paragraph for sectPr that is inside a paragraph
            let parent = sectPr.parentNode;
            while (parent && parent.nodeName !== "w:p") {
                parent = parent.parentNode;
            }
            
            // If sectPr is inside a paragraph, it's a section break
            if (parent) {
                // Calculate position
                let position = 0;
                const paragraphs = docXml.getElementsByTagNameNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "p");
                for (let j = 0; j < paragraphs.length; j++) {
                    if (paragraphs[j] === parent) {
                        position = j;
                        break;
                    }
                }
                
                breaks.push({
                    type: "section",
                    position: position,
                    element: sectPr
                });
            }
        }
        
        return breaks;
    }

    /**
     * Render HTML content with pagination
     */
    function renderPaginatedContent(htmlContent) {
        return new Promise((resolve, reject) => {
            try {
                // Create container for the content
                const container = document.createElement('div');
                container.style.position = 'absolute';
                container.style.left = '-9999px';
                container.style.width = `${pageSettings.width - pageSettings.margins.left - pageSettings.margins.right}px`;
                container.style.fontFamily = 'Arial, sans-serif';
                container.style.fontSize = '12pt';
                container.style.lineHeight = '1.15';
                container.innerHTML = htmlContent;
                document.body.appendChild(container);
                
                // Find all explicit page breaks in the HTML
                const explicitBreakElements = container.querySelectorAll('hr.page-break');
                const breakPoints = Array.from(explicitBreakElements).map(el => {
                    return {
                        element: el,
                        type: 'explicit'
                    };
                });
                
                // Calculate content height to estimate implicit breaks
                const contentHeight = container.scrollHeight;
                const contentPerPage = pageSettings.height - pageSettings.margins.top - pageSettings.margins.bottom;
                const estimatedPageCount = Math.max(1, Math.ceil(contentHeight / contentPerPage));
                
                console.log(`Content height: ${contentHeight}px, available height per page: ${contentPerPage}px`);
                console.log(`Estimated page count: ${estimatedPageCount}`);
                
                // Split content into pages
                const pages = [];
                
                // If no explicit breaks, render as one page
                if (breakPoints.length === 0) {
                    // Create pages based on content height
                    const totalPages = Math.max(1, Math.ceil(contentHeight / contentPerPage));
                    
                    for (let i = 0; i < totalPages; i++) {
                        // Clone the container for each page
                        const pageDiv = container.cloneNode(true);
                        // Set height to content per page
                        pageDiv.style.height = `${contentPerPage}px`;
                        pageDiv.style.overflow = 'hidden';
                        
                        // Position to show only current page's content
                        pageDiv.style.position = 'relative';
                        pageDiv.style.top = `-${i * contentPerPage}px`;
                        pageDiv.style.clip = `rect(${i * contentPerPage}px, ${pageSettings.width}px, ${(i + 1) * contentPerPage}px, 0)`;
                        
                        // Render page to canvas
                        renderPageToCanvas(pageDiv, i).then(pageCanvas => {
                            pages[i] = pageCanvas;
                            
                            // Check if all pages are rendered
                            if (pages.filter(Boolean).length === totalPages) {
                                documentPages = pages;
                                document.body.removeChild(container);
                                resolve();
                            }
                        }).catch(reject);
                    }
                } else {
                    // Handle explicit page breaks
                    let currentPos = 0;
                    let pageIndex = 0;
                    
                    // Function to render a page from start to end position
                    const renderPageFromRange = (startPos, endPos, index) => {
                        const pageDiv = container.cloneNode(true);
                        
                        // Hide all elements outside the range
                        let allElements = Array.from(pageDiv.querySelectorAll('*'));
                        let visibleElements = allElements.slice(startPos, endPos);
                        
                        allElements.forEach(el => {
                            if (!visibleElements.includes(el)) {
                                el.style.display = 'none';
                            }
                        });
                        
                        renderPageToCanvas(pageDiv, index).then(pageCanvas => {
                            pages[index] = pageCanvas;
                            
                            // Check if all pages are rendered
                            if (pages.filter(Boolean).length === breakPoints.length + 1) {
                                documentPages = pages;
                                document.body.removeChild(container);
                                resolve();
                            }
                        }).catch(reject);
                    };
                    
                    // Process each break point to create pages
                    breakPoints.forEach((breakPoint, i) => {
                        // Find position of the break element
                        const elements = Array.from(container.querySelectorAll('*'));
                        const breakPos = elements.indexOf(breakPoint.element);
                        
                        // Render page from current position to break point
                        renderPageFromRange(currentPos, breakPos, pageIndex);
                        
                        // Update current position and page index
                        currentPos = breakPos + 1; // Skip the break element
                        pageIndex++;
                        
                        // If last break point, render the final page
                        if (i === breakPoints.length - 1) {
                            renderPageFromRange(currentPos, elements.length, pageIndex);
                        }
                    });
                }
            } catch (error) {
                reject(error);
            }
        });
    }
    
    /**
     * Render a div to canvas
     */
    function renderPageToCanvas(div, pageIndex) {
        return new Promise((resolve, reject) => {
            try {
                document.body.appendChild(div);
                
                // Use html2canvas to render the div to a canvas
                html2canvas(div, {
                    width: pageSettings.width,
                    height: pageSettings.height,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: '#ffffff'
                }).then(function(renderedCanvas) {
                    // Create a new canvas with proper dimensions including margins
                    const pageCanvas = document.createElement('canvas');
                    pageCanvas.width = pageSettings.width;
                    pageCanvas.height = pageSettings.height;
                    
                    const ctx = pageCanvas.getContext('2d');
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
                    
                    // Draw the rendered content with margins
                    ctx.drawImage(
                        renderedCanvas, 
                        0, 0, renderedCanvas.width, renderedCanvas.height,
                        pageSettings.margins.left, pageSettings.margins.top, 
                        pageSettings.width - pageSettings.margins.left - pageSettings.margins.right, 
                        pageSettings.height - pageSettings.margins.top - pageSettings.margins.bottom
                    );
                    
                    document.body.removeChild(div);
                    resolve(pageCanvas);
                }).catch(function(error) {
                    document.body.removeChild(div);
                    reject(error);
                });
            } catch (error) {
                if (document.body.contains(div)) {
                    document.body.removeChild(div);
                }
                reject(error);
            }
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
        if (!documentLoaded) {
            const error = new Error("Document not loaded");
            console.error(error);
            if (onCompletion) onCompletion(error);
            return;
        }

        if (index < 0 || index >= documentPages.length) {
            const error = new Error("Invalid page index");
            console.error(error);
            if (onCompletion) onCompletion(error);
            return;
        }

        currPageIndex = index;
        
        // Get the page canvas for the current index
        const pageCanvas = documentPages[index];
        
        if (!pageCanvas) {
            const error = new Error("Page not rendered");
            console.error(error);
            if (onCompletion) onCompletion(error);
            return;
        }
        
        // Draw the page to the main canvas with scaling and rotation
        const ctx = canvas.getContext('2d');
        
        // Clear the canvas
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        
        // Set canvas dimensions
        canvas.width = pageCanvas.width * scale;
        canvas.height = pageCanvas.height * scale;
        
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

    /**
     * Alternative implementation using docx-preview library
     * This can be used as a fallback if mammoth.js doesn't properly handle pagination
     */
    function renderWithDocxPreview() {
        return new Promise((resolve, reject) => {
            try {
                // Create a temporary container
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
                    breakPages: true,
                    renderHeaders: true,
                    renderFooters: true,
                    renderFootnotes: true
                }).then(result => {
                    // Find all page divs
                    const pageElements = tempContainer.querySelectorAll('.docx-page');
                    const pages = [];
                    
                    // Process each page
                    let processedPages = 0;
                    pageElements.forEach((pageElement, index) => {
                        // Get page dimensions
                        const rect = pageElement.getBoundingClientRect();
                        
                        // Use html2canvas to render the page
                        html2canvas(pageElement, {
                            width: rect.width,
                            height: rect.height,
                            useCORS: true,
                            allowTaint: true,
                            backgroundColor: '#ffffff'
                        }).then(pageCanvas => {
                            pages[index] = pageCanvas;
                            processedPages++;
                            
                            // Check if all pages are processed
                            if (processedPages === pageElements.length) {
                                document.body.removeChild(tempContainer);
                                documentPages = pages;
                                resolve();
                            }
                        }).catch(error => {
                            console.error("Error rendering page:", error);
                            processedPages++;
                            
                            // Continue even with errors
                            if (processedPages === pageElements.length) {
                                document.body.removeChild(tempContainer);
                                documentPages = pages.filter(Boolean);
                                resolve();
                            }
                        });
                    });
                }).catch(reject);
            } catch (error) {
                reject(error);
            }
        });
    }
}
