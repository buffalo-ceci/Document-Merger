import { renderAsync } from 'docx-preview';
import { jsPDF } from 'jspdf';
import { PDFDocument } from 'pdf-lib';
import html2canvas from 'html2canvas';

// jsPDF needs html2canvas to be available globally in some environments
if (typeof window !== 'undefined') {
  (window as any).html2canvas = html2canvas;
}

export async function convertDocxToPdf(file: File): Promise<ArrayBuffer> {
  const iframe = document.createElement('iframe');

  // Let docx-preview use the document's own page dimensions by isolating in an iframe
  iframe.style.position = 'absolute';
  iframe.style.left = '-9999px';
  iframe.style.top = '-9999px';
  iframe.style.width = '1200px';
  iframe.style.height = '1200px';
  document.body.appendChild(iframe);

  try {
    const iframeDoc = iframe.contentDocument || iframe.contentWindow?.document;
    if (!iframeDoc) throw new Error("Could not access iframe document");

    iframeDoc.open();
    iframeDoc.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { margin: 0; padding: 0; background: #e0e0e0; color: #000; font-family: "Microsoft JhengHei", "PMingLiU", "MingLiU", "DFKai-SB", sans-serif; }
            table { border-collapse: collapse; }
            .docx-wrapper { padding: 0 !important; }
            section.docx { box-shadow: none !important; margin: 0 !important; }
          </style>
        </head>
        <body>
          <div id="docx-container"></div>
        </body>
      </html>
    `);
    iframeDoc.close();

    const container = iframeDoc.getElementById('docx-container') as HTMLElement;

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
      debug: false, // enables additional logging
      renderHeaders: true, // render headers
      renderFooters: true, // render footers
      renderFootnotes: true,
      renderEndnotes: true,
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

    const pdfDoc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    const pdfWidth = pdfDoc.internal.pageSize.getWidth();
    const pdfHeight = pdfDoc.internal.pageSize.getHeight();

    for (let i = 0; i < elementsToRender.length; i++) {
      const el = elementsToRender[i] as HTMLElement;

      // Ensure the section has a white background for the PDF
      const originalBg = el.style.background || el.style.backgroundColor;
      const originalBoxShadow = el.style.boxShadow;
      const originalMargin = el.style.margin;

      el.style.backgroundColor = '#ffffff';
      el.style.boxShadow = 'none';
      el.style.margin = '0';

      const canvas = await html2canvas(el, {
        scale: 2, // Better quality
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff',
        // Make sure html2canvas uses the isolated iframe window to parse CSS properly
        window: iframe.contentWindow as unknown as Window
      } as any);

      // Restore background just in case
      el.style.background = originalBg;
      el.style.boxShadow = originalBoxShadow;
      el.style.margin = originalMargin;

      const imgData = canvas.toDataURL('image/jpeg', 0.95);
      const imgProps = pdfDoc.getImageProperties(imgData);
      const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;

      if (i > 0) {
        pdfDoc.addPage();
      }

      let heightLeft = imgHeight;
      let position = 0;

      pdfDoc.addImage(imgData, 'JPEG', 0, position, pdfWidth, imgHeight);
      heightLeft -= pdfHeight;

      // If the section is somehow taller than A4, slice it (fallback)
      while (heightLeft > 0) {
        position = heightLeft - imgHeight;
        pdfDoc.addPage();
        pdfDoc.addImage(imgData, 'JPEG', 0, position, pdfWidth, imgHeight);
        heightLeft -= pdfHeight;
      }
    }

    return pdfDoc.output('arraybuffer');
  } finally {
    if (document.body.contains(iframe)) {
      document.body.removeChild(iframe);
    }
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
