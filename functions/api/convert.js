import { PDFDocument } from 'pdf-lib';

export async function onRequest(context) {
  // 处理 CORS
  if (context.request.method === 'OPTIONS') {
    return new Response(null, {
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      },
    });
  }

  if (context.request.method !== 'POST') {
    return new Response('Method not allowed', { status: 405 });
  }

  try {
    const formData = await context.request.formData();
    const file = formData.get('file');
    const conversionType = formData.get('type') || 'pdf-to-word';

    if (!file) {
      return new Response('No file uploaded', { status: 400 });
    }

    const fileBuffer = await file.arrayBuffer();
    
    if (conversionType === 'pdf-to-word') {
      // 使用 pdf-lib 提取文本内容
      const pdfDoc = await PDFDocument.load(fileBuffer);
      const pages = pdfDoc.getPages();
      
      // 提取所有页面的文本
      let textContent = '';
      for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const { width, height } = page.getSize();
        
        // 这里需要添加文本提取逻辑
        // 由于 pdf-lib 的限制，我们需要使用其他库来提取文本
        // 建议使用 pdf.js 或 pdf-parse 等库
        
        textContent += `Page ${i + 1}\n`;
      }
      
      // 创建 Word 文档内容
      const docxContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r>
                <w:t>${textContent}</w:t>
              </w:r>
            </w:p>
          </w:body>
        </w:document>`;
      
      // 返回转换后的文档
      return new Response(docxContent, {
        headers: {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="converted.docx"`,
          'Access-Control-Allow-Origin': '*',
        },
      });
    }

    return new Response('Unsupported conversion type', { status: 400 });
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*',
      },
    });
  }
} 
