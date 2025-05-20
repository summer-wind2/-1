from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import requests
import io
import logging
import sys
import zipfile

app = Flask(__name__)
CORS(app)

# 日志配置
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

ILOVEPDF_PUBLIC_KEY = 'project_public_91561bf682ce7ae31cab1e9b9176cf16_kE7Tub838515d440b1a86535ae57b031227e1'  # 请替换为你自己的public_key

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if 'file' not in request.files:
            return jsonify({'error': '请选择要转换的PDF文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '请选择要转换的PDF文件'}), 400

        file.stream.seek(0)

        # 1. 创建任务
        start_resp = requests.post(
            'https://api.ilovepdf.com/v1/start',
            json={'public_key': ILOVEPDF_PUBLIC_KEY, 'tool': 'pdf2word'}
        )
        start_json = start_resp.json()
        if 'task' not in start_json:
            return jsonify({'error': f"iLovePDF创建任务失败: {start_json}"})
        task = start_json['task']

        # 2. 上传文件
        upload_resp = requests.post(
            'https://api.ilovepdf.com/v1/upload',
            files={'file': (file.filename, file.stream, 'application/pdf')},
            data={'task': task}
        )
        upload_json = upload_resp.json()
        if 'server_filename' not in upload_json:
            return jsonify({'error': f"iLovePDF上传文件失败: {upload_json}"})
        server_filename = upload_json['server_filename']

        # 3. 处理任务
        process_resp = requests.post(
            'https://api.ilovepdf.com/v1/process',
            json={'task': task}
        )
        process_json = process_resp.json()
        if 'status' not in process_json or process_json['status'] != 'TaskSuccess':
            return jsonify({'error': f"iLovePDF处理任务失败: {process_json}"})

        # 4. 下载结果
        download_resp = requests.get(
            f'https://api.ilovepdf.com/v1/download/{task}',
            stream=True
        )
        if download_resp.status_code != 200:
            return jsonify({'error': 'iLovePDF下载文件失败'}), 500

        # 5. 解压zip，提取docx
        with zipfile.ZipFile(io.BytesIO(download_resp.content)) as zf:
            print('zip包内容:', zf.namelist())
            docx_bytes = None
            for name in zf.namelist():
                if name.endswith('.docx'):
                    docx_bytes = zf.read(name)
                    print('docx文件大小:', len(docx_bytes))
                    break
            if not docx_bytes:
                return jsonify({'error': 'zip包中未找到docx文件'}), 500

        # 6. 返回Word文件
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=file.filename.rsplit('.', 1)[0] + '.docx'
        )

    except Exception as e:
        logger.error(f"文件处理过程中出错: {str(e)}")
        return jsonify({'error': f'文件转换失败: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000) 
