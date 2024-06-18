const express = require('express');
const path = require('path');
const app = express();
const PORT = 5000;

// 假设这是最新的软件版本号
const latestVersion = "0.1.7";

// 设置一个路由来提供最新版本信息
app.get('/latest_version.json', (req, res) => {
    res.json({ version: latestVersion });
});

// 设置一个路由来托管软件下载文件（如果需要）
app.get('/software_latest.exe', (req, res) => {
    // 这里可以添加代码来处理文件的托管逻辑
    // 例如，从文件系统读取文件并发送给客户端
    res.sendFile('setup.exe');
});

// 设置静态文件目录
app.use(express.static(path.join(__dirname, 'public')));

// 路由处理函数
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public/index.html'));
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`服务器正在端口上运行 ${PORT}`);
});

