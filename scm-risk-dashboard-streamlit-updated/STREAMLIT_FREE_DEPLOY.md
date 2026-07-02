# 免费部署方案：Streamlit Community Cloud

## 适合做什么

Streamlit Community Cloud 适合先把 SCM 看板免费跑起来：

- 项目组可以通过一个网址访问。
- 你可以在网页里上传最新版 Follow Up。
- 上传后系统会重新计算并展示看板。

## 需要注意

免费平台更适合试运行，不适合作为最终长期数据仓库：

- 上传的文件保存在应用运行环境里。
- 如果免费应用休眠、重启或重建，可能需要重新上传 Follow Up。
- 如果后续要稳定保存历史版本、操作日志、多人权限，建议再升级到云服务器或接数据库/云盘。

## 部署前准备

1. 准备一个 GitHub 仓库。
2. 把这个项目代码推到 GitHub。
3. 登录 Streamlit Community Cloud。
4. 新建 App，入口文件选择：

```text
streamlit_app.py
```

5. 在 Streamlit 的 Secrets 里设置共享密码：

```toml
SCM_DASHBOARD_PASSWORD = "这里改成你的共享密码"
```

## 部署后怎么用

1. 打开 Streamlit 给你的网址。
2. 输入共享密码。
3. 左侧上传最新版 Follow Up。
4. 系统重新计算后显示看板。
5. 其他人打开同一个网址即可查看。

## GitHub Pages 为什么不适合

GitHub Pages 只能稳定托管静态页面，不适合处理：

- 上传 Excel
- 后台重新计算
- 保存最新版本

所以 GitHub Pages 可以做“只读展示版”，但不适合你现在这个需要上传和计算的工作流。
