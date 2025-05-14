function DocxHandler(canvas) {
    var self = this;
    self.DPI = 96;

    var docxFile;
    var currPageIndex = 0;
    var pageContents = []; // Will store HTML content for each page
    var renderedPages = []; // Rendered canvases for each page
    var docxUrl; // Store URL for debugging

    // Standard page dimensions - US Letter in pixels at 96 DPI
    var DEFAULT_WIDTH = 816; // 8.5" × 96dpi = 816px
    var DEFAULT_HEIGHT = 1056; // 11" × 96dpi = 1056px
    var DEFAULT_MARGIN = 96; // 1" margin
    var CONTENT_HEIGHT = DEFAULT_HEIGHT - (DEFAULT_MARGIN * 2); // Available content height per page
    var MAX_IMAGES_PER_PAGE = 5; // Limit images per page to ensure better pagination

    this.loadDocument = function(documentUrl, onCompletion, onError) {
        benchmark.time("Document loaded");
        docxUrl = documentUrl; // Store for debugging
        
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

    this.processDocument = function () {
        return new Promise((resolve, reject) => {
            try {
                // Process document using mammoth
                processWithMammoth().then(resolve).catch(reject);
            } catch (error) {
                console.error("Error in processDocument:", error);
                reject(error);
            }
        });
    };

    // Process document using mammoth
    function processWithMammoth() {
        return new Promise((resolve, reject) => {
            try {
                if (typeof mammoth === 'undefined') {
                    return reject(new Error("mammoth.js not available"));
                }

                console.log("Processing document with mammoth.js");
                mammoth.convertToHtml(
                    { arrayBuffer: docxFile },
                    { 
                        styleMap: [
                            "w:br[type='page'] => hr.page-break",
                            "p[style-name='heading 1'] => h1:fresh",
                            "p[style-name='heading 2'] => h2:fresh",
                            "p[style-name='heading 3'] => h3:fresh",
                            "table => table",
                            "tr => tr",
                            "tc => td"
                        ]
                    }
                ).then(function (result) {
                    const html = result.value;
                    const warnings = result.messages;
                    
                    if (warnings.length > 0) {
                        console.warn("Mammoth warnings:", warnings);
                    }
                    
                    // Create a temp div to hold the full content
                    const tempDiv = document.createElement('div');
                    tempDiv.innerHTML = html;
                    tempDiv.style.position = 'absolute';
                    tempDiv.style.left = '-9999px';
                    tempDiv.style.width = (DEFAULT_WIDTH - DEFAULT_MARGIN * 2) + 'px';
                    document.body.appendChild(tempDiv);
                    
                    // First check for explicit page breaks
                    const pageBreaks = tempDiv.querySelectorAll('hr.page-break');
                    console.log(`Found ${pageBreaks.length} page breaks in document`);
                    
                    if (pageBreaks.length > 0) {
                        // We have explicit page breaks - use them
                        createPagesFromExplicitBreaks(tempDiv, pageBreaks);
                    } else {
                        // No explicit breaks, create intelligent pagination
                        createAutomaticPagination(tempDiv);
                    }
                    
                    // Clean up
                    document.body.removeChild(tempDiv);
                    console.log(`Document processed into ${pageContents.length} pages`);
                    resolve();
                }).catch(function (error) {
                    console.error("Mammoth conversion error:", error);
                    reject(error);
                });
            } catch (error) {
                console.error("Error in processWithMammoth:", error);
                reject(error);
            }
        });
    }
    
    // Create pages based on explicit page breaks
    function createPagesFromExplicitBreaks(contentDiv, pageBreaks) {
        pageContents = [];
        
        // Split by page breaks
        const breakElements = Array.from(pageBreaks);
        breakElements.unshift(null); // Add null as first break point (start of document)
        
        // Process each segment between breaks
        for (let i = 0; i < breakElements.length; i++) {
            const start = breakElements[i]; // Current break (or null for start)
            const end = (i < breakElements.length - 1) ? breakElements[i + 1] : null; // Next break or null
            
            // Create a container for this page content
            const pageContent = document.createElement('div');
            
            // Get all elements between breaks
            let currentNode = start ? start.nextSibling : contentDiv.firstChild;
            
            while (currentNode && currentNode !== end) {
                const nextNode = currentNode.nextSibling;
                pageContent.appendChild(currentNode.cloneNode(true));
                currentNode = nextNode;
            }
            
            // Further divide page if it has too many images
            const pageImages = pageContent.querySelectorAll('img');
            if (pageImages.length > MAX_IMAGES_PER_PAGE) {
                // Too many images, split into sub-pages
                const subPages = splitContentByImageCount(pageContent);
                subPages.forEach(subPage => {
                    pageContents.push(wrapInPageContainer(subPage.innerHTML));
                });
            } else {
                // Page is acceptable, add it
                pageContents.push(wrapInPageContainer(pageContent.innerHTML));
            }
        }
    }
    
    // Create intelligent pagination when no explicit breaks exist
    function createAutomaticPagination(contentDiv) {
        pageContents = [];
        
        // First, collect all elements
        const allElements = Array.from(contentDiv.childNodes);
        
        // Group elements into pages
        let currentPage = document.createElement('div');
        let imageCounter = 0;
        let estimatedHeight = 0;
        
        allElements.forEach(element => {
            // Clone the element to manipulate it
            const clonedElement = element.cloneNode(true);
            
            // Check if this is an image
            const isImage = element.nodeName === 'IMG' || 
                           (element.nodeName === 'P' && element.querySelector('img'));
            
            if (isImage) {
                imageCounter++;
            }
            
            // Estimate element height (very rough approximation)
            let elementHeight = 0;
            if (isImage) {
                elementHeight = 200; // Base image height estimate
            } else if (element.nodeName === 'TABLE') {
                const rows = element.querySelectorAll('tr').length;
                elementHeight = rows * 30 + 50; // Rough estimate for table height
            } else if (element.nodeName === 'H1') {
                elementHeight = 60;
            } else if (element.nodeName === 'H2') {
                elementHeight = 45;
            } else if (element.nodeName === 'P') {
                // Approximate paragraph height based on text length
                const text = element.textContent || '';
                const lines = Math.ceil(text.length / 80); // ~80 chars per line
                elementHeight = lines * 20 + 10; // 20px per line + margin
            } else {
                elementHeight = 30; // Default height for other elements
            }
            
            estimatedHeight += elementHeight;
            
            // Check if we should start a new page
            if ((imageCounter > MAX_IMAGES_PER_PAGE) || (estimatedHeight > CONTENT_HEIGHT)) {
                // Current page is full, add it to pages and start new page
                if (currentPage.childNodes.length > 0) {
                    pageContents.push(wrapInPageContainer(currentPage.innerHTML));
                    currentPage = document.createElement('div');
                    imageCounter = isImage ? 1 : 0;
                    estimatedHeight = elementHeight;
                }
            }
            
            // Add element to current page
            currentPage.appendChild(clonedElement);
        });
        
        // Add the last page if it has content
        if (currentPage.childNodes.length > 0) {
            pageContents.push(wrapInPageContainer(currentPage.innerHTML));
        }
    }
    
    // Split content based on image count
    function splitContentByImageCount(content) {
        const subPages = [];
        let currentPage = document.createElement('div');
        let imageCounter = 0;
        
        Array.from(content.childNodes).forEach(node => {
            const isImage = node.nodeName === 'IMG' || 
                          (node.nodeName === 'P' && node.querySelector('img'));
            
            if (isImage) {
                imageCounter++;
                
                // Check if we should start a new page
                if (imageCounter > MAX_IMAGES_PER_PAGE && currentPage.childNodes.length > 0) {
                    subPages.push(currentPage);
                    currentPage = document.createElement('div');
                    imageCounter = 1; // Reset counter (starting with this image)
                }
            }
            
            // Add node to current page
            currentPage.appendChild(node.cloneNode(true));
        });
        
        // Add the last page if it has content
        if (currentPage.childNodes.length > 0) {
            subPages.push(currentPage);
        }
        
        return subPages;
    }

    // Wrap content in a page container with styles
    function wrapInPageContainer(content) {
        return `
            <div class="docx-page" style="
                width: ${DEFAULT_WIDTH}px;
                min-height: ${DEFAULT_HEIGHT}px;
                padding: ${DEFAULT_MARGIN}px;
                background-color: white;
                box-sizing: border-box;
                font-family: Arial, sans-serif;
                font-size: 12pt;
                line-height: 1.5;
                position: relative;
                margin: 0 auto;
                overflow: hidden;
            ">
                ${content}
            </div>
        `;
    }

    this.pageCount = function () {
        return pageContents.length || 1;
    };

    this.originalWidth = function () {
        return DEFAULT_WIDTH;
    };

    this.originalHeight = function () {
        return DEFAULT_HEIGHT;
    };

    this.drawDocument = function (scale, rotation, onCompletion) {
        benchmark.time("Document drawn");
        self.redraw(scale, rotation, currPageIndex, function (err) {
            benchmark.timeEnd("Document drawn");
            if (onCompletion) onCompletion(err);
        });
    };

    this.redraw = function (scale, rotation, index, onCompletion) {
        try {
            if (pageContents.length === 0) {
                console.error("No pages available to render");
                if (onCompletion) onCompletion(new Error("No pages available"));
                return;
            }

            // Ensure valid index
            index = Math.max(0, Math.min(index, pageContents.length - 1));
            currPageIndex = index;

            console.log(`Rendering page ${index + 1} of ${pageContents.length}`);

            // Use cached page if available
            if (renderedPages[index]) {
                console.log(`Using cached page ${index + 1}`);
                const ctx = canvas.getContext('2d');
                canvas.width = renderedPages[index].width * scale;
                canvas.height = renderedPages[index].height * scale;
                ctx.scale(scale, scale);
                ctx.drawImage(renderedPages[index], 0, 0);
                if (onCompletion) onCompletion();
                return;
            }

            // Create temporary container
            const tempContainer = document.createElement('div');
            tempContainer.innerHTML = pageContents[index];
            tempContainer.style.position = 'absolute';
            tempContainer.style.left = '-9999px';
            tempContainer.style.top = '0';
            document.body.appendChild(tempContainer);

            // Style for better rendering
            const pageElement = tempContainer.querySelector('.docx-page');
            if (!pageElement) {
                console.error("No .docx-page element found in page content");
                document.body.removeChild(tempContainer);
                if (onCompletion) onCompletion(new Error("Invalid page structure"));
                return;
            }

            // Style all elements
            stylePageContent(pageElement);

            // Find all images and wait for them to load
            const images = Array.from(pageElement.querySelectorAll('img'));
            console.log(`Page ${index + 1} has ${images.length} images`);

            const imagePromises = images.map(img => {
                return new Promise((resolve) => {
                    img.style.maxWidth = '100%'; // Ensure images fit in page
                    
                    if (img.complete) {
                        resolve();
                    } else {
                        img.onload = () => resolve();
                        img.onerror = () => {
                            console.warn(`Image failed to load: ${img.src}`);
                            resolve(); // Resolve anyway to continue
                        };
                    }
                });
            });

            Promise.all(imagePromises).then(() => {
                // Render the page to canvas
                return html2canvas(pageElement, {
                    width: DEFAULT_WIDTH,
                    height: DEFAULT_HEIGHT,
                    useCORS: true,
                    allowTaint: true,
                    backgroundColor: 'white',
                    scale: 1,
                    logging: false
                });
            }).then(renderedCanvas => {
                // Store the rendered canvas
                renderedPages[index] = renderedCanvas;

                // Draw to the output canvas
                const ctx = canvas.getContext('2d');
                canvas.width = renderedCanvas.width * scale;
                canvas.height = renderedCanvas.height * scale;

                // Apply rotation if needed
                if (rotation) {
                    ctx.save();
                    ctx.translate(canvas.width / 2, canvas.height / 2);
                    ctx.rotate(rotation * Math.PI / 180);
                    ctx.scale(scale, scale);
                    ctx.drawImage(renderedCanvas, -renderedCanvas.width / 2, -renderedCanvas.height / 2);
                    ctx.restore();
                } else {
                    ctx.scale(scale, scale);
                    ctx.drawImage(renderedCanvas, 0, 0);
                }

                // Clean up
                document.body.removeChild(tempContainer);
                console.log(`Page ${index + 1} rendered successfully`);
                if (onCompletion) onCompletion();
            }).catch(error => {
                console.error(`Error rendering page ${index + 1}:`, error);
                
                // Create a fallback error page
                const errorCanvas = document.createElement('canvas');
                errorCanvas.width = DEFAULT_WIDTH;
                errorCanvas.height = DEFAULT_HEIGHT;
                const ctx = errorCanvas.getContext('2d');
                
                // White background
                ctx.fillStyle = 'white';
                ctx.fillRect(0, 0, DEFAULT_WIDTH, DEFAULT_HEIGHT);
                
                // Error message
                ctx.fillStyle = 'red';
                ctx.font = '16px Arial';
                ctx.fillText(`Error rendering page ${index + 1}`, 50, 50);
                ctx.fillText(error.message, 50, 80);
                ctx.fillText(`Document URL: ${docxUrl}`, 50, 110);
                
                // Store and display
                renderedPages[index] = errorCanvas;
                
                const canvasCtx = canvas.getContext('2d');
                canvas.width = DEFAULT_WIDTH * scale;
                canvas.height = DEFAULT_HEIGHT * scale;
                canvasCtx.scale(scale, scale);
                canvasCtx.drawImage(errorCanvas, 0, 0);
                
                // Clean up
                if (tempContainer.parentNode) {
                    document.body.removeChild(tempContainer);
                }
                
                if (onCompletion) onCompletion(error);
            });
        } catch (error) {
            console.error("Unexpected error in redraw:", error);
            if (onCompletion) onCompletion(error);
        }
    };

    // Style page content for better rendering
    function stylePageContent(pageElement) {
        try {
            // Style paragraphs
            const paragraphs = pageElement.querySelectorAll('p');
            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                p.style.margin = '0 0 10px 0';
                p.style.color = '#000000';
            }
            
            // Style headings
            const headings = pageElement.querySelectorAll('h1, h2, h3, h4, h5, h6');
            for (let i = 0; i < headings.length; i++) {
                const h = headings[i];
                h.style.margin = '15px 0 10px 0';
                h.style.color = '#000000';
            }
            
            // Style tables
            const tables = pageElement.querySelectorAll('table');
            for (let i = 0; i < tables.length; i++) {
                const table = tables[i];
                table.style.borderCollapse = 'collapse';
                table.style.width = '100%';
                table.style.margin = '10px 0';
                
                const cells = table.querySelectorAll('td, th');
                for (let j = 0; j < cells.length; j++) {
                    const cell = cells[j];
                    cell.style.border = '1px solid #ccc';
                    cell.style.padding = '5px';
                    cell.style.color = '#000000';
                }
            }
            
            // Style images
            const images = pageElement.querySelectorAll('img');
            for (let i = 0; i < images.length; i++) {
                const img = images[i];
                img.style.maxWidth = '100%';
                img.style.height = 'auto';
                img.setAttribute('crossOrigin', 'anonymous');
            }
            
            // Style lists
            const lists = pageElement.querySelectorAll('ul, ol');
            for (let i = 0; i < lists.length; i++) {
                const list = lists[i];
                list.style.margin = '10px 0';
                list.style.paddingLeft = '20px';
                list.style.color = '#000000';
            }
        } catch (error) {
            console.error("Error in stylePageContent:", error);
        }
    }

    this.applyToCanvas = function (apply) {
        apply(canvas);
    };

    this.createCanvases = function (callback, dpiCalcFunction, fromPage, pageCount) {
        const canvases = [];
        
        if (pageContents.length === 0) {
            console.error("No pages available for createCanvases");
            callback([]);
            return;
        }
        
        const toPage = Math.min(pageContents.length - 1, fromPage + pageCount - 1);
        let completed = 0;

        for (let i = fromPage; i <= toPage; i++) {
            this.redraw(1, 0, i, function () {
                canvases.push({
                    canvas: canvas.cloneNode(true),
                    originalDocumentDpi: self.DPI
                });
                
                completed++;
                if (completed === (toPage - fromPage + 1)) {
                    callback(canvases);
                }
            });
        }
    };
}
