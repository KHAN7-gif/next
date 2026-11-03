import {createWorker} from 'tesseract.js';
import {getDocument,GlobalWorkerOptions} from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import {Document,Packer,Paragraph,PageBreak} from 'docx';
GlobalWorkerOptions.workerSrc=workerUrl;
export async function convertScannedPdf(file,onProgress=()=>{}){
 const pdf=await getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
 let worker;const paragraphs=[];let recognized=0;
 try{
  onProgress('Memuatkan enjin OCR (English)...');
  worker=await createWorker('eng',1,{logger:m=>{if(m.status==='recognizing text')onProgress('Mengecam teks: '+Math.round((m.progress||0)*100)+'%')}});
  for(let n=1;n<=pdf.numPages;n++){
   onProgress('OCR halaman '+n+' / '+pdf.numPages);
   const page=await pdf.getPage(n);
   const viewport=page.getViewport({scale:Math.min(2,1400/page.getViewport({scale:1}).width)});
   const canvas=document.createElement('canvas');
   canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
   const context=canvas.getContext('2d',{alpha:false});
   await page.render({canvasContext:context,viewport,background:'white'}).promise;
   const result=await worker.recognize(canvas);
   canvas.width=0;canvas.height=0;
   const text=result.data.text||'';
   recognized+=text.trim().length;
   for(const line of text.split(/\r?\n/))paragraphs.push(new Paragraph({text:line.trim(),spacing:{after:line.trim()?0:120}}));
   if(n<pdf.numPages)paragraphs.push(new Paragraph({children:[new PageBreak()]}));
  }
 }finally{if(worker)await worker.terminate();await pdf.destroy()}
 if(!recognized)throw Error('OCR tidak menemui teks pada dokumen ini.');
 onProgress('Menjana dokumen Word...');
 const doc=new Document({sections:[{children:paragraphs.length?paragraphs:[new Paragraph('')]}]});
 const blob=await Packer.toBlob(doc);
 const url=URL.createObjectURL(blob),a=document.createElement('a');
 a.href=url;a.download=file.name.replace(/\.pdf$/i,'')+'-OCR.docx';a.click();
 setTimeout(()=>URL.revokeObjectURL(url),30000);
 return {pages:paragraphs.length};
}
