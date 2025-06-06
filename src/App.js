import React, { useState } from 'react';
import { useDropzone } from 'react-dropzone';
import {
  Box,
  Button,
  Container,
  Typography,
  CircularProgress,
  Alert,
  Snackbar,
  LinearProgress,
} from '@mui/material';
import axios from 'axios';

// API URL配置
const API_URL = 'https://1-chi-dusky.vercel.app//api/convert';
const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20MB

function App() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState(false);
  const [progress, setProgress] = useState(0);

  const onDrop = (acceptedFiles) => {
    const selectedFile = acceptedFiles[0];
    if (!selectedFile) {
      setError('请选择文件');
      return;
    }
    if (selectedFile.type !== 'application/pdf') {
      setError('请上传PDF文件');
      return;
    }
    if (selectedFile.size > MAX_FILE_SIZE) {
      setError('文件大小不能超过20MB');
      return;
    }
    setFile(selectedFile);
    setError('');
    setProgress(0);
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/pdf': ['.pdf'],
    },
    multiple: false,
    maxSize: MAX_FILE_SIZE,
  });

  const handleConvert = async () => {
    if (!file) {
      setError('请先选择文件');
      return;
    }

    setLoading(true);
    setError('');
    setSuccess(false);
    setProgress(0);

    try {
      const formData = new FormData();
      formData.append('file', file);

      const response = await axios.post(API_URL, formData, {
        responseType: 'blob',
        timeout: 60000, // 增加超时时间到60秒
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        onUploadProgress: (progressEvent) => {
          const percentCompleted = Math.round((progressEvent.loaded * 100) / progressEvent.total);
          setProgress(percentCompleted);
        },
      });

      // 检查响应状态
      if (response.status === 200) {
        // 创建下载链接并触发下载
        const blob = new Blob([response.data], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = file.name.replace('.pdf', '.docx');
        document.body.appendChild(link);
        link.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(link);
        setSuccess(true);
        setFile(null);
        setProgress(100);
      } else {
        throw new Error('转换失败');
      }
    } catch (err) {
      console.error('转换错误:', err);
      
      if (err.response) {
        // 服务器返回错误
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const errorText = reader.result;
            const error = JSON.parse(errorText);
            setError(error.error || '转换失败，请重试');
          } catch (e) {
            setError('转换失败，请重试');
          }
        };
        reader.readAsText(err.response.data);
      } else if (err.request) {
        // 请求发送失败
        setError('服务器无响应，请检查网络连接');
      } else {
        // 其他错误
        setError('转换失败，请重试');
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <Container maxWidth="sm">
      <Box
        sx={{
          marginTop: 8,
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          gap: 2,
        }}
      >
        <Typography variant="h4" component="h1" gutterBottom>
          PDF转Word转换器
        </Typography>

        <Box
          {...getRootProps()}
          sx={{
            width: '100%',
            height: 200,
            border: '2px dashed #ccc',
            borderRadius: 2,
            display: 'flex',
            flexDirection: 'column',
            justifyContent: 'center',
            alignItems: 'center',
            cursor: 'pointer',
            backgroundColor: isDragActive ? '#f0f0f0' : 'white',
            '&:hover': {
              borderColor: '#1976d2',
              backgroundColor: '#f5f5f5',
            },
            transition: 'all 0.3s ease',
          }}
        >
          <input {...getInputProps()} />
          <Typography textAlign="center" gutterBottom>
            {isDragActive ? '将文件拖放到这里' : '点击或拖放PDF文件到这里'}
          </Typography>
          <Typography variant="body2" color="textSecondary">
            最大20MB
          </Typography>
        </Box>

        {file && (
          <Typography variant="body1">
            已选择文件: {file.name}
          </Typography>
        )}

        {progress > 0 && progress < 100 && (
          <Box sx={{ width: '100%', mt: 2 }}>
            <LinearProgress variant="determinate" value={progress} />
          </Box>
        )}

        {error && (
          <Alert severity="error" sx={{ width: '100%' }}>
            {error}
          </Alert>
        )}

        <Button
          variant="contained"
          onClick={handleConvert}
          disabled={!file || loading}
          sx={{ mt: 2 }}
        >
          {loading ? (
            <CircularProgress size={24} color="inherit" />
          ) : (
            '开始转换'
          )}
        </Button>

        <Snackbar
          open={success}
          autoHideDuration={6000}
          onClose={() => setSuccess(false)}
        >
          <Alert severity="success" sx={{ width: '100%' }}>
            转换成功！
          </Alert>
        </Snackbar>
      </Box>
    </Container>
  );
}

export default App; 
