# PDF转Word在线转换器

这是一个简单的在线PDF转Word转换工具，使用Flask作为后端，React作为前端。

## 功能特点

- 拖放式文件上传
- 实时文件转换
- 自动下载转换后的文件
- 支持所有主流操作系统
- 美观的用户界面

## 安装说明

### 后端设置

1. 确保已安装Python 3.7+
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```
3. 运行后端服务器：
   ```bash
   python app.py
   ```

### 前端设置

1. 确保已安装Node.js 14+
2. 进入前端目录：
   ```bash
   cd frontend
   ```
3. 安装依赖：
   ```bash
   npm install
   ```
4. 运行开发服务器：
   ```bash
   npm start
   ```

## 部署说明

### 域名设置

要通过4renpk.my或www.4renpk.my访问，需要：

1. 在域名管理面板中添加A记录，指向您的服务器IP
2. 配置nginx反向代理：

```nginx
server {
    listen 80;
    server_name 4renpk.my www.4renpk.my;

    location / {
        proxy_pass http://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
    }

    location /api {
        proxy_pass http://localhost:5000;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection 'upgrade';
        proxy_set_header Host $host;
        proxy_cache_bypass $http_upgrade;
    }
}
```

## 生产环境部署

1. 构建前端：
   ```bash
   cd frontend
   npm run build
   ```

2. 使用gunicorn运行后端：
   ```bash
   gunicorn -w 4 -b 0.0.0.0:5000 app:app
   ```

## 注意事项

- 确保服务器有足够的存储空间
- 定期清理临时文件
- 建议设置文件大小限制
- 配置SSL证书以支持HTTPS 