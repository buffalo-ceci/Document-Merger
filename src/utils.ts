import { renderAsync } from 'docx-preview';
import { jsPDF } from 'jspdf';
import { PDFDocument } from 'pdf-lib';
import html2canvas from 'html2canvas';

// jsPDF needs html2canvas to be available globally in some environments
if (typeof window !== 'undefined') {
  (window as any).html2canvas = html2canvas;
}

export async function convertDocxToPdf(file: File): Promise<ArrayBuffer> {
  const container = document.createElement('div');
  
  // We use a fixed width that corresponds to A4 at 96 DPI (approx 794px)
  // docx-preview will try to render pages inside this.
  container.style.width = '800px'; 
  container.style.position = 'absolute';
  container.style.left = '0';
  container.style.top = '0';
  container.style.zIndex = '-9999'; // Hide behind other elements
  container.style.background = '#e0e0e0'; // docx-preview usually renders white pages on a gray background
  
  document.body.appendChild(container);

  try {
    const arrayBuffer = await file.arrayBuffer();
    
    await renderAsync(arrayBuffer, container, null, {
      className: 'docx', // class name/prefix for default and mathml elements
      inWrapper: true, // enables rendering of wrapper around document content
      ignoreWidth: false, // disables rendering width of page
      ignoreHeight: false, // disables rendering height of page
      ignoreFonts: false, // disables fonts rendering
      breakPages: true, // enables page breaking on page breaks
      ignoreLastRenderedPageBreak: true, // disables page breaking on lastRenderedPageBreak elements
      experimental: true, // enables experimental features (tab stops calculation)
      trimXmlDeclaration: true, // if true, xml declaration will be removed from xml string before parsing
      useBase64URL: true, // if true, images, fonts, etc. will be converted to base 64 URL, otherwise URL.createObjectURL is used
      useMathMLPolyfill: false, // includes MathML polyfills for chrome, edge, etc.
      showChanges: false, // experimental: show tracked changes
      debug: false, // enables additional logging
    });

    // Remove unsupported images like WMF/EMF which cause html2canvas to crash
    const images = container.querySelectorAll('img');
    images.forEach(img => {
      if (
        img.src.startsWith('data:image/x-wmf') || 
        img.src.startsWith('data:image/wmf') || 
        img.src.startsWith('data:image/x-emf') ||
        img.src.startsWith('data:image/emf')
      ) {
        const placeholder = document.createElement('div');
        placeholder.style.border = '1px dashed #ccc';
        placeholder.style.padding = '10px';
        placeholder.style.color = '#999';
        placeholder.style.fontSize = '10px';
        placeholder.style.display = 'inline-block';
        placeholder.innerText = '[Unsupported Image Format]';
        img.parentNode?.replaceChild(placeholder, img);
      }
    });

    // docx-preview renders pages as <section class="docx"> inside the container.
    // We can capture each section individually to preserve pagination perfectly.
    const sections = Array.from(container.querySelectorAll('section.docx'));
    const elementsToRender = sections.length > 0 ? sections : [container];

    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    const pdfWidth = doc.internal.pageSize.getWidth();
    const pdfHeight = doc.internal.pageSize.getHeight();

    for (let i = 0; i < elementsToRender.length; i++) {
      const el = elementsToRender[i] as HTMLElement;
      
      // Ensure the section has a white background for the PDF
      const originalBg = el.style.background || el.style.backgroundColor;
      el.style.backgroundColor = '#ffffff';

      const canvas = await html2canvas(el, {
        scale: 2, // Better quality
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      });

      // Restore background just in case
      el.style.background = originalBg;

      const imgData = canvas.toDataURL('image/jpeg', 0.95);
      const imgProps = doc.getImageProperties(imgData);
      const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;
      
      if (i > 0) {
        doc.addPage();
      }
      
      let heightLeft = imgHeight;
      let position = 0;

      doc.addImage(imgData, 'JPEG', 0, position, pdfWidth, imgHeight);
      heightLeft -= pdfHeight;

      // If the section is somehow taller than A4, slice it (fallback)
      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        doc.addPage();
        doc.addImage(imgData, 'JPEG', 0, position, pdfWidth, imgHeight);
        heightLeft -= pdfHeight;
      }
    }

    return doc.output('arraybuffer');
  } finally {
    document.body.removeChild(container);
  }
}

export async function mergeFiles(files: File[]): Promise<Uint8Array> {
  const mergedPdf = await PDFDocument.create();

  for (const file of files) {
    let pdfBuffer: ArrayBuffer;
    
    if (file.name.toLowerCase().endsWith('.pdf')) {
      pdfBuffer = await file.arrayBuffer();
    } else if (file.name.toLowerCase().endsWith('.docx')) {
      pdfBuffer = await convertDocxToPdf(file);
    } else {
      throw new Error(`Unsupported file type: ${file.name}`);
    }

    const pdf = await PDFDocument.load(pdfBuffer);
    const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    copiedPages.forEach((page) => mergedPdf.addPage(page));
  }

  return await mergedPdf.save();
}
