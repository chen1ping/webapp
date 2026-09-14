# Render + Supabase 免费云端部署指南

## 架构概览

```
用户浏览器 → Render (Flask + Gunicorn) → Supabase (PostgreSQL)
                           ↓
                    Supabase Storage
              (上传文件、生成文件都存这里)
```

全部免费，不需要信用卡。

---

## 第一步：注册 Supabase（数据库 + 文件存储）

1. 打开 https://supabase.com/
2. 点击 "Start your project"，用 GitHub 账号登录（免信用卡）
3. 点击 "New Project" 创建新项目
   - **Name**: 随便填，比如 `webapp`
   - **Database Password**: 点 "Generate a password" 生成一个，**保存好后面要用**
   - **Region**: 选离你近的，比如 `Southeast Asia (Singapore)`
   - 点击 "Create new project"
4. 等 1-2 分钟，项目创建完成

**获取数据库连接串：**
5. 进入项目 → 左侧菜单 **Settings** → **Database**
6. 找到 **Connection string** → **URI**，复制那串 `postgresql://postgres:xxx@xxx.supabase.co:5432/postgres`
7. 保存好，后面填到 Render 环境变量里

**创建存储桶：**
8. 左侧菜单 **Storage** → 点击 "New bucket"
9. Name 填 `webapp-files`，勾选 "Make bucket public"（公开访问）
10. 点击 "Create bucket"

**获取 Supabase URL 和 Key：**
11. 左侧菜单 **Project Settings** → **API**
12. 复制 **Project URL**（形如 `https://xxx.supabase.co`）
13. 复制 **anon public** key（不是 service_role key）

---

## 第二步：注册 Render（应用服务器）

1. 打开 https://render.com/
2. 点击 "Get Started"，用 GitHub 账号登录（免信用卡）
3. 免费层每月有 750 小时运行时间，够用

---

## 第三步：把代码上传到 GitHub

Render 需要从 GitHub 拉取代码。

1. 打开 https://github.com/ 注册/登录
2. 点击右上角 **+** → **New repository**
   - Repository name: 随便填，比如 `webapp`
   - 选 **Public** 或 **Private** 都行
   - 点击 "Create repository"
3. 把本地 `C:\WebApp` 里的代码上传到这个仓库

> **提示**：如果你不会用 Git，可以直接在 GitHub 网页上点 "Add file" → "Upload files"，把 `C:\WebApp` 里的所有文件（除了 `data/`、`__pycache__/`、`.env`）拖拽上传。

**必须上传的文件：**
- `app.py`
- `config.py`
- `database.py`
- `db_filter.py`
- `storage.py`
- `models.py`
- `step1_generator.py`
- `step23_processor.py`
- `requirements.txt`
- `render.yaml`
- `templates/` 文件夹（整个上传）
- `static/` 文件夹（整个上传）

---

## 第四步：在 Render 上部署

1. 登录 Render 后，点击 **New** → **Web Service**
2. 选择你刚才创建的 GitHub 仓库
3. 填写配置：
   - **Name**: 随便填，比如 `my-webapp`（会成为你的子域名）
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt && python -c "from database import init_db_tables; init_db_tables()"`
   - **Start Command**: `gunicorn -w 2 -b 0.0.0.0:$PORT app:app`
   - **Instance Type**: 选 **Free**
4. 点击 **Advanced** → **Add Environment Variable**，添加以下变量：

| Key | Value | 说明 |
|-----|-------|------|
| `DB_TYPE` | `postgres` | 数据库类型 |
| `DATABASE_URL` | 第一步复制的 PostgreSQL 连接串 | 数据库地址 |
| `STORAGE_TYPE` | `supabase` | 文件存储类型 |
| `SUPABASE_URL` | 第一步复制的 Project URL | Supabase 地址 |
| `SUPABASE_KEY` | 第一步复制的 anon key | Supabase 公钥 |
| `SUPABASE_BUCKET` | `webapp-files` | 存储桶名称 |
| `SECRET_KEY` | 随机一串字符 | 会话加密用 |
| `PORT` | `10000` | Render 用的端口 |
| `PYTHON_VERSION` | `3.11.0` | Python 版本 |

5. 点击 **Create Web Service**
6. 等待部署完成（第一次可能需要 2-5 分钟）
7. 部署成功后，顶部会显示你的访问地址，形如 `https://my-webapp.onrender.com`

---

## 第五步：初始化管理员账号

部署成功后，第一次访问需要创建管理员账号。

1. 打开你的 Render 地址
2. 点击注册，创建第一个账号（自动成为管理员）
3. 登录后就可以正常使用了

---

## 数据迁移（可选）

如果你想把本地的数据搬到云端：

1. 把本地 `data/system.db` 文件上传到 Render 服务器（或用 SCP）
2. 运行迁移脚本：
   ```bash
   python migrate_sqlite_to_mysql.py data/system.db
   ```
   （PostgreSQL 模式下脚本会自动适配）

或者直接在云端从零开始，重新导入数据。

---

## 常见问题

**Q: Render 免费版会休眠吗？**
A: 会，15 分钟没人访问就休眠。下次访问需要等 10-30 秒启动。免费版每月 750 小时，够全天跑。

**Q: Supabase 免费额度够吗？**
A: 免费层 500MB 数据库 + 1GB 存储 + 每月 5GB 流量，几十万条企业数据完全够用。

**Q: 访问速度怎么样？**
A: Supabase 新加坡节点 + Render 美国节点，国内访问可能稍慢（1-3秒）。如果追求速度可以用国内云。

**Q: 想升级配置怎么办？**
A: Render 和 Supabase 后台都可以随时升级付费方案，数据不用迁移。
