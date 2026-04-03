# PPT复刻Skill - 完整部署指南

本文档包含从飞书应用创建到Skill上线的完整步骤。

---

## 一、飞书开发者后台配置

### 1.1 创建飞书应用

1. 访问 [飞书开放平台](https://open.feishu.cn/)
2. 登录后点击 **"创建企业自建应用"**
3. 填写应用信息：
   - **应用名称**：PPT复刻助手
   - **应用描述**：智能复刻PPT模板，一键生成可编辑的新PPT
   - **应用图标**：上传一个PPT相关的图标
4. 点击 **"创建应用"**

### 1.2 配置应用凭证

1. 进入应用详情页 → **"凭证与基础信息"**
2. 记录以下信息（后续部署需要）：
   - **App ID**（例如：`cli_xxxxxxxxxxxxxxxx`）
   - **App Secret**（点击显示并复制）
3. **安全设置**：
   - 添加 **IP白名单**（如果使用云服务器，添加服务器公网IP）
   - 或勾选 **"无需IP白名单"**（开发测试阶段）

### 1.3 开通所需权限

进入 **"权限管理"** 页面，开通以下权限：

| 权限 | 权限代码 | 用途 |
|------|----------|------|
| 查看文档元数据 | `drive:drive:readonly` | 获取PPT文件信息 |
| 下载文件 | `drive:file:download` | 下载用户分享的PPT |
| 上传文件 | `drive:file:upload` | 上传生成的PPT |
| 创建分享链接 | `drive:file:share` | 生成可分享的PPT链接 |
| 读取用户基本信息 | `contact:user.base:readonly` | 获取发送者名称 |
| 发送消息 | `im:message:send` | 发送处理结果 |
| 接收消息 | `im:message:receive` | 接收用户消息 |

**操作步骤**：
1. 在权限管理页面搜索上述权限代码
2. 点击每个权限右侧的 **"申请"** 按钮
3. 填写申请理由：
   > "PPT复刻助手需要此权限以实现：读取用户分享的PPT模板、下载文件进行解析、上传生成的新PPT文件、发送处理结果给用户。"
4. 提交申请，等待审核（通常几分钟到几小时）

### 1.4 配置事件订阅

进入 **"事件与回调"** → **"事件订阅"** 页面：

1. **加密密钥（Encrypt Key）**：
   - 点击 **"重新生成"**
   - 复制并保存（后续部署需要）

2. **验证令牌（Verification Token）**：
   - 点击 **"重新生成"**
   - 复制并保存（后续部署需要）

3. **请求地址（URL）**：
   - 暂时填写：`https://your-domain.com/webhook`
   - 部署完成后再更新为真实地址

4. **订阅事件类型**：
   - 点击 **"添加事件"**
   - 搜索并添加：`im.message.receive_v1`（接收消息事件）

### 1.5 配置机器人能力

进入 **"机器人"** 页面：

1. 打开 **"启用机器人"** 开关
2. 配置机器人信息：
   - **显示名称**：PPT复刻助手
   - **描述**：发送飞书PPT链接，我帮你复刻模板
   - **头像**：上传图标
3. 配置 **"消息卡片"** 能力：
   - 打开 **"支持消息卡片"** 开关

### 1.6 发布应用

进入 **"版本管理与发布"** 页面：

1. 点击 **"创建版本"**
2. 填写版本信息：
   - **版本号**：1.0.0
   - **更新说明**：初始版本，支持PPT模板复刻
3. 点击 **"保存"** → **"申请发布"**
4. 通知企业管理员审核通过

---

## 二、Work Buddy Skill配置

### 2.1 创建Skill

1. 打开 Work Buddy 应用
2. 进入 **"技能中心"** → **"自定义技能"**
3. 点击 **"创建技能"**
4. 填写技能信息：

**基础信息**：
- **技能名称**：PPT复刻助手
- **技能ID**：`ppt-clone-assistant`
- **描述**：智能复刻PPT模板，一键生成可编辑的新PPT
- **图标**：上传PPT相关图标

**触发方式**：
- **触发类型**：消息触发
- **触发词**：见下文"触发词配置"

**运行方式**：
- **执行方式**：Webhook回调
- **回调地址**：`https://your-domain.com/workbuddy/skill`

### 2.2 配置环境变量

在Skill设置中配置以下环境变量：

```bash
# 飞书应用凭证（必填）
FEISHU_APP_ID=cli_xxxxxxxxxxxxxxxx
FEISHU_APP_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

# 飞书事件订阅配置（必填）
FEISHU_ENCRYPT_KEY=xxxxxxxxxxxxxxxx
FEISHU_VERIFICATION_TOKEN=xxxxxxxx

# 应用配置（可选）
DOWNLOAD_DIR=./downloads
OUTPUT_DIR=./output
DEFAULT_SLIDE_COUNT=10
```

---

## 三、部署方案

### 方案A：Railway部署（推荐，免费额度足够）

**优点**：
- 免费额度：每月500小时运行时间 + 5GB流量
- 自动HTTPS
- 一键部署
- 内置环境变量管理

**部署步骤**：

1. **准备代码**：
   ```bash
   # 确保代码已推送到GitHub
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/yourusername/ppt-clone-skill.git
   git push -u origin main
   ```

2. **创建Railway项目**：
   - 访问 [Railway](https://railway.app/)
   - 使用GitHub账号登录
   - 点击 **"New Project"**
   - 选择 **"Deploy from GitHub repo"**
   - 选择你的代码仓库

3. **配置环境变量**：
   - 进入项目 → **"Variables"**
   - 添加上述所有环境变量

4. **配置启动命令**：
   - 进入 **"Settings"**
   - 设置 **"Start Command"**：
     ```bash
     python main.py --mode webhook --port $PORT
     ```

5. **获取域名**：
   - Railway会自动分配域名
   - 复制域名（例如：`ppt-clone.up.railway.app`）

6. **更新飞书配置**：
   - 将Railway域名更新到飞书事件订阅URL：
     ```
     https://ppt-clone.up.railway.app/webhook
     ```

7. **更新Work Buddy配置**：
   - Skill回调地址：
     ```
     https://ppt-clone.up.railway.app/workbuddy/skill
     ```

### 方案B：阿里云函数计算（FC）

**优点**：
- 按调用次数计费，低频使用成本极低
- 自动扩缩容
- 国内访问速度快

**部署步骤**：

1. **准备代码**：
   创建 `index.py`：
   ```python
   import json
   from main import handle_webhook, handle_workbuddy_skill
   
   def handler(event, context):
       evt = json.loads(event)
       path = evt.get('path', '')
       
       if path == '/webhook':
           return handle_webhook(evt)
       elif path == '/workbuddy/skill':
           return handle_workbuddy_skill(evt)
       
       return {'statusCode': 404, 'body': 'Not Found'}
   ```

2. **创建函数**：
   - 登录 [阿里云函数计算控制台](https://fc.console.aliyun.com/)
   - 创建服务和函数
   - 运行时：Python 3.9
   - 上传代码ZIP包

3. **配置触发器**：
   - 添加HTTP触发器
   - 认证方式：匿名访问
   - 获取公网访问地址

4. **配置环境变量**：
   - 在函数配置中添加所有环境变量

5. **更新配置**：
   - 将函数计算域名更新到飞书和Work Buddy

### 方案C：自有服务器（VPS）

**适用场景**：已有服务器资源

**部署步骤**：

1. **服务器准备**：
   ```bash
   # 安装Python 3.9+
   sudo apt update
   sudo apt install python3.9 python3.9-pip
   
   # 安装Git
   sudo apt install git
   ```

2. **部署代码**：
   ```bash
   # 克隆代码
   git clone https://github.com/yourusername/ppt-clone-skill.git
   cd ppt-clone-skill
   
   # 创建虚拟环境
   python3.9 -m venv venv
   source venv/bin/activate
   
   # 安装依赖
   pip install -r requirements.txt
   ```

3. **配置环境变量**：
   ```bash
   # 创建.env文件
   cat > .env << EOF
   FEISHU_APP_ID=cli_xxxxxxxxxxxxxxxx
   FEISHU_APP_SECRET=xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   FEISHU_ENCRYPT_KEY=xxxxxxxxxxxxxxxx
   FEISHU_VERIFICATION_TOKEN=xxxxxxxx
   EOF
   ```

4. **使用Gunicorn运行**：
   ```bash
   # 安装gunicorn
   pip install gunicorn
   
   # 启动服务
   gunicorn -w 2 -b 0.0.0.0:8000 "main:create_app()"
   ```

5. **配置Nginx反向代理**：
   ```nginx
   server {
       listen 80;
       server_name your-domain.com;
       
       location / {
           proxy_pass http://127.0.0.1:8000;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
   }
   ```

6. **配置SSL（Let's Encrypt）**：
   ```bash
   sudo apt install certbot python3-certbot-nginx
   sudo certbot --nginx -d your-domain.com
   ```

7. **使用Systemd守护进程**：
   创建 `/etc/systemd/system/ppt-clone.service`：
   ```ini
   [Unit]
   Description=PPT Clone Skill
   After=network.target

   [Service]
   User=www-data
   WorkingDirectory=/path/to/ppt-clone-skill
   Environment="PATH=/path/to/ppt-clone-skill/venv/bin"
   ExecStart=/path/to/ppt-clone-skill/venv/bin/gunicorn -w 2 -b 127.0.0.1:8000 "main:create_app()"
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```

   启动服务：
   ```bash
   sudo systemctl enable ppt-clone
   sudo systemctl start ppt-clone
   ```

---

## 四、触发词与Prompt配置

### 4.1 Work Buddy触发词配置

在Work Buddy技能设置中添加以下触发词：

**精确匹配**：
```
复刻PPT
PPT复刻
克隆PPT
PPT克隆
复制PPT模板
```

**正则匹配**：
```regex
(复刻|克隆|复制).{0,5}PPT
PPT.{0,5}(复刻|克隆|复制)
帮我.{0,10}PPT.{0,10}模板
```

**包含关键词**：
```
feishu.cn/slides
飞书PPT
PPT链接
```

### 4.2 系统Prompt配置

在Work Buddy技能设置中配置系统Prompt：

```
你是PPT复刻助手，专门帮助用户复刻PPT模板。

## 你的能力
1. 接收用户分享的飞书PPT链接
2. 解析PPT模板的样式特征（配色、字体、版式）
3. 生成一个新的、可编辑的PPT文件
4. 返回飞书PPT链接给用户

## 支持的参数
用户可以在消息中指定以下参数：
- 页面数："15页"、"--pages 15"、"-p 15"
- 主色调："蓝色主题"、"--color 0000FF"、"-c FF5500"
- 保留动画："保留动画"、"--keep-animation"

## 处理流程
1. 提取用户消息中的飞书PPT链接
2. 解析用户指定的参数（如有）
3. 调用PPT复刻服务进行处理
4. 向用户返回处理结果和新PPT链接

## 回复格式
处理成功时：
"✅ PPT复刻完成！\n\n📄 原文件：[文件名]\n📊 页面数：[N] 页\n⏱️ 耗时：[N] 秒\n\n🔗 复刻后的PPT：[链接]"

处理失败时：
"❌ PPT复刻失败\n\n错误信息：[错误详情]\n\n💡 建议：[解决方案]"

## 注意事项
- 只处理飞书PPT链接（feishu.cn/slides/xxx）
- 如果用户没有指定页面数，默认生成10页
- 如果用户没有指定颜色，使用原模板配色
- 处理时间通常在10-30秒之间，请提示用户耐心等待
```

### 4.3 示例对话配置

在Work Buddy中添加示例对话，帮助用户了解如何使用：

**示例1 - 基础使用**：
- **用户**：复刻这个PPT https://xxx.feishu.cn/slides/xxx
- **助手**：收到！正在为您复刻PPT模板，请稍候...
- **助手**：✅ PPT复刻完成！...

**示例2 - 指定页面数**：
- **用户**：帮我复刻这个模板，要20页 https://xxx.feishu.cn/slides/xxx
- **助手**：收到！正在为您复刻PPT模板（20页），请稍候...

**示例3 - 指定颜色**：
- **用户**：复刻PPT，蓝色主题，15页 https://xxx.feishu.cn/slides/xxx
- **助手**：收到！正在为您复刻PPT模板（15页，蓝色主题），请稍候...

---

## 五、验证测试

### 5.1 本地测试

```bash
# 1. 设置环境变量
export FEISHU_APP_ID=cli_xxx
export FEISHU_APP_SECRET=xxx

# 2. 运行本地测试
python main.py --mode local --url "https://xxx.feishu.cn/slides/xxx" --pages 15
```

### 5.2 Webhook测试

```bash
# 1. 启动本地服务器
python main.py --mode webhook --port 8080

# 2. 使用ngrok暴露公网地址
ngrok http 8080

# 3. 更新飞书事件订阅URL为ngrok地址
# 4. 在飞书中发送PPT链接测试
```

### 5.3 健康检查

部署完成后，访问健康检查端点：
```
https://your-domain.com/health
```

应返回：
```json
{
  "status": "ok",
  "timestamp": "2026-04-03T17:00:00"
}
```

---

## 六、常见问题排查

### 6.1 飞书事件订阅验证失败

**问题**：飞书提示 "请求地址校验失败"

**解决**：
1. 检查URL是否正确（需要https）
2. 检查Verification Token是否一致
3. 检查服务是否正常响应

### 6.2 权限不足错误

**问题**：提示 "权限不足" 或 "Permission denied"

**解决**：
1. 检查飞书应用是否已开通所需权限
2. 检查权限申请是否已通过审核
3. 重新发布应用版本

### 6.3 PPT下载失败

**问题**：提示 "文件下载失败"

**解决**：
1. 检查PPT链接是否有效
2. 检查文件是否已被删除
3. 检查应用是否有下载权限

### 6.4 处理超时

**问题**：Work Buddy提示处理超时

**解决**：
1. 检查服务器性能是否足够
2. 考虑使用异步处理模式
3. 增加超时时间配置

---

## 七、后续优化建议

1. **添加缓存**：缓存已解析的模板，避免重复下载
2. **支持更多格式**：扩展支持Keynote、Google Slides等
3. **批量处理**：支持一次处理多个PPT链接
4. **模板市场**：建立常用模板库，快速复刻热门模板
5. **用户配置**：支持用户保存常用配置（默认页面数、默认颜色等）

---

**部署完成！** 🎉

现在用户可以在飞书对话中@你的机器人，发送PPT链接进行复刻了。
