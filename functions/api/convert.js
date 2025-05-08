import { PDFDocument } from 'pdf-lib';
import { convert } from 'pdf2docx';

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
      // 使用 pdf2docx 进行转换
      const docxBuffer = await convert(fileBuffer);
      
      // 返回转换后的文档
      return new Response(docxBuffer, {
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
