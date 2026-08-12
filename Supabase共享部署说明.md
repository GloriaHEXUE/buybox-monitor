# Supabase 共享部署说明

更新时间：2026-08-12

这份文档用于说明 `buybox-monitor` 如何从“仅浏览器本地缓存”升级为“多人访问同一网址时共享同一份数据”的版本。

当前项目已经完成以下能力：

- 网站代码已接入 Supabase 共享读写
- Vercel 已可通过环境变量连接 Supabase
- Keepa、映射信息、SKU / ASIN 监控清单、在线新增监控都可写入共享数据
- Keepa 趋势历史已支持共享保存
- ASIN 主图已支持共享保存

---

## 1. 当前共享架构

当前线上版本的数据职责如下：

- `GitHub`
  - 存放网站代码
  - 代码更新后触发 Vercel 部署
- `Vercel`
  - 托管正式网站
  - 网址为：
    - [https://buybox-monitor-beryl.vercel.app/](https://buybox-monitor-beryl.vercel.app/)
- `Supabase`
  - 存放网站共享数据
  - 多人打开同一个 Vercel 网址时，读取的是同一份 Supabase 数据

也就是说：

- 改页面、改功能、改逻辑：走 GitHub
- 上传 Keepa、更新映射、更新监控清单、在线新增监控：写入 Supabase

---

## 2. 当前共享数据保存方式

当前项目采用的是“整站状态统一保存”的方案。

Supabase 中使用一张表：

- 表名：`app_state`

表内保存一条主记录：

- `state_key = global`

这一条记录的 `payload` 中保存整站当前状态，包括：

- `keepaRows`
- `yesterdayKeepaRows`
- `mappingRows`
- `monitorRows`
- `onlineRows`
- `sourceReports`
- `history`
- `todayBuyBox`
- `yesterdayBuyBox`
- `keepaUploadReports`

优点：

- 改造快
- 对现有单页应用侵入小
- 非常适合当前“内部共用一个运营工具”的场景

当前阶段不需要拆成多张业务表，就可以实现共享。

---

## 3. Supabase 建表方法

### 3.1 新建项目

在 Supabase 中新建一个项目。

当前实际使用的项目地址是：

- `https://zvbiifbyardvgwptqmol.supabase.co`

### 3.2 执行建表 SQL

进入：

- `SQL Editor`

执行以下 SQL：

```sql
create table if not exists public.app_state (
  state_key text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);
```

建表完成后，`app_state` 就会用于保存整站共享数据。

---

## 4. Supabase 权限配置方法

当前项目为了让正式网页可直接读写共享数据，使用的是匿名读写策略。

进入：

- `SQL Editor`

执行以下 SQL：

```sql
alter table public.app_state enable row level security;

create policy "allow read app_state"
on public.app_state
for select
to anon
using (true);

create policy "allow write app_state"
on public.app_state
for insert
to anon
with check (true);

create policy "allow update app_state"
on public.app_state
for update
to anon
using (true)
with check (true);
```

说明：

- 当前策略适合你们内部共用网站的场景
- 这意味着匿名访问网页的人也可以写入这张共享状态表
- 如果以后对外开放，建议改成登录后读写，或者细化权限

---

## 5. 需要的 Supabase 项目信息

进入 Supabase 项目后台后，需要拿到：

- `Project URL`
- `Publishable key`

注意：

- 使用的是 `Publishable key`
- 不是 `Secret key`

---

## 6. Vercel 环境变量配置方法

进入 Vercel 项目的：

- `Settings`
- `Environment Variables`

新增以下 4 个环境变量：

```text
VITE_SUPABASE_URL=https://zvbiifbyardvgwptqmol.supabase.co
VITE_SUPABASE_ANON_KEY=你的 Supabase Publishable key
VITE_SUPABASE_STATE_TABLE=app_state
VITE_SUPABASE_STATE_KEY=global
```

推荐环境：

- `Production`
- `Preview`

添加完成后，需要重新部署。

---

## 7. Vercel 重新部署方法

进入：

- `Vercel`
- `Deployments`

找到最新一条部署，点击：

- `Redeploy`

或者当 GitHub `main` 分支有新提交时，Vercel 也会自动部署。

判断部署是否成功：

- 状态显示 `Ready`

---

## 8. 当前代码中的共享实现位置

当前共享逻辑主要在以下文件：

- [C:\Users\Lenovo\Documents\buy\buybox-monitor\src\lib\sharedState.ts](C:\Users\Lenovo\Documents\buy\buybox-monitor\src\lib\sharedState.ts)
- [C:\Users\Lenovo\Documents\buy\buybox-monitor\src\App.tsx](C:\Users\Lenovo\Documents\buy\buybox-monitor\src\App.tsx)

说明：

- `sharedState.ts`
  - 负责连接 Supabase
  - 负责读取 / 写入 `app_state`
- `App.tsx`
  - 负责在页面启动时加载共享数据
  - 负责本地修改后自动回写共享数据

---

## 9. 当前共享后的实际行为

当网页已正确接入 Supabase 后，以下数据会共享：

- Keepa 数据
- 映射信息
- SKU / ASIN 监控清单
- 在线新增监控 ASIN
- Buy Box 丢失 / 恢复状态
- 价格趋势与排名趋势
- ASIN 主图

也就是说：

- 你在正式网址上传 Keepa
- 或修改映射
- 或新增在线监控

其他人打开同一个网址，也会看到相同内容。

---

## 10. Keepa 上传规则说明

当前 Keepa 采用“昨日 / 今日双槽位”逻辑：

### 10.1 每日标准上传方式

例如：

- 2026-08-12
  - 昨日 = 2026-08-11 的 Keepa
  - 今日 = 2026-08-12 的 Keepa

- 2026-08-13
  - 昨日 = 2026-08-12 的 Keepa
  - 今日 = 2026-08-13 的 Keepa

### 10.2 覆盖规则

上传为“昨日数据”时：

- 会覆盖旧的昨日 Keepa 基准
- 会清空当前今日结果，等待你重新上传今日数据

上传为“今日数据”时：

- 会覆盖旧的今日 Keepa
- 会与当前保存的昨日 Keepa 对比，重新生成当天结果

### 10.3 趋势保留规则

虽然“昨日 / 今日槽位”会滚动覆盖，但趋势历史会保留最近 5 天。

当前版本已经修复：

- 趋势历史可共享保存
- 刷新后不会只剩当天

---

## 11. 为什么之前无痕页看不到 ASIN 图片

之前的问题根因是：

- Keepa 图片字段在共享保存时被清空

所以：

- 本地页因为仍然保留当前内存数据，所以能看到图
- 无痕页 / 别人打开网页时，从 Supabase 读取的共享数据里没有图，所以看不到

当前版本已修复：

- Keepa 图片字段会跟随共享数据一起保存

注意：

- 修复上线前已经写入 Supabase 的旧 Keepa 数据，图片字段不会自动补回来
- 需要重新上传一次 Keepa，主图才会重新写入共享数据

---

## 12. 如何确认共享是否生效

推荐用以下 3 个方法验证：

### 方法一：无痕窗口验证

1. 在正式网址中上传一份数据
2. 等几秒
3. 打开无痕窗口访问同一网址
4. 看数据是否同步出现

### 方法二：Supabase 表验证

进入：

- `Table Editor`
- `app_state`

查看：

- `updated_at`

如果你刚在网页上修改过数据，而这里的时间也更新了，说明共享写入成功。

### 方法三：数据更新页验证

当前 `数据更新` 页面已增加：

- `趋势天数`
- `云端同步`

可直接查看：

- 当前趋势已累计几天
- 最近一次成功同步到云端的时间

---

## 13. 本地开发配置方法

如果需要本地联调共享版本，可在项目根目录创建：

- `.env.local`

内容示例：

```text
VITE_SUPABASE_URL=https://zvbiifbyardvgwptqmol.supabase.co
VITE_SUPABASE_ANON_KEY=你的 Supabase Publishable key
VITE_SUPABASE_STATE_TABLE=app_state
VITE_SUPABASE_STATE_KEY=global
```

然后运行：

```bash
npm install
npm run dev
```

---

## 14. 后续如需继续升级的方向

当前方案已经可以满足内部共享使用，但如果以后需要更强能力，可以再考虑第二阶段升级：

1. 把整站 `app_state` 拆成多张业务表
2. 增加用户登录与权限控制
3. 增加上传操作日志
4. 增加历史版本回滚
5. 增加多人并发编辑保护

但在当前阶段，不建议为了“更标准”而过早复杂化。

---

## 15. 给后续 Codex 或同事的一句话

如果后续要继续维护，请先理解：

- 当前正式共享数据保存在 Supabase `app_state` 表
- Vercel 通过环境变量连接 Supabase
- Keepa 采用“昨日 / 今日”双槽位 + 最近 5 天趋势历史
- 映射、监控清单、在线新增、主图、趋势都已经接入共享保存

