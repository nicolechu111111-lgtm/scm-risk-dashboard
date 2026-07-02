# SCM 在线看板部署说明

这个版本把现有 SCM 看板变成一个可部署的网页服务：

- 项目组打开同一个网址查看同一版看板。
- 你上传最新版 Follow Up 后，后台自动重新计算。
- 重新计算完成后，其他人刷新页面就能看到最新数据。
- 可设置共享密码，避免外部人员访问。

## 本地试运行

在项目目录运行：

```bash
SCM_DASHBOARD_PASSWORD='设置一个共享密码' PORT=8080 python webapp/scm_web_dashboard.py
```

然后打开：

```text
http://127.0.0.1:8080/dashboard
```

第一次进入时，如果还没有上传 Follow Up，会显示上传页面。上传后会自动生成看板。

## Docker 部署

构建镜像：

```bash
docker build -t scm-dashboard .
```

启动服务：

```bash
docker run -d \
  --name scm-dashboard \
  -p 8080:8080 \
  -e SCM_DASHBOARD_PASSWORD='设置一个共享密码' \
  -v scm-dashboard-data:/data \
  scm-dashboard
```

访问：

```text
http://服务器IP:8080/dashboard
```

## 数据保存位置

在线版会把上传的 Follow Up 和生成后的看板数据保存在：

```text
/data
```

如果用 Docker 部署，请务必挂载 volume，否则容器重建后上传历史会丢失。

## 推荐后续部署方式

如果只是项目组内部使用，推荐两种：

1. 公司内网服务器或一台固定电脑  
   优点：最快落地，文件不出公司网络。

2. 云服务器加密码访问  
   优点：美国团队也能随时打开；后续可以加账号权限、操作日志和自动邮件提醒。

当前版本已经适合作为“第一版在线网页看板”。后续如果多人同时编辑或需要权限分层，再升级成数据库版本。
