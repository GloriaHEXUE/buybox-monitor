# Supabase 共享部署说明

这份说明用于把当前 `buybox-monitor` 从“浏览器本地缓存版”升级为“全员共享数据版”。

升级目标：

- 你上传 Keepa、映射信息、监控清单后
- 数据写入 Supabase 云端
- 所有人打开同一个 Vercel 链接，都能看到同一份最新数据

---

## 1. 当前代码状态

项目已经接入了 Supabase 共享状态读写骨架，相关文件：

- [src/lib/sharedState.ts](C:\Users\Lenovo\Documents\buy\buybox-monitor\src\lib\sharedState.ts)
- [src/App.tsx](C:\Users\Lenovo\Documents\buy\buybox-monitor\src\App.tsx)

当前逻辑：

1. 如果未配置 Supabase 环境变量  
   项目继续走本地缓存逻辑

2. 如果已配置 Supabase 环境变量  
   项目会：
   - 启动时尝试从 Supabase 拉取共享状态
   - 页面数据变化后自动回写 Supabase

也就是说，现在离“共享版正式可用”只差：

- 建 Supabase 项目
- 建表
- 配环境变量
- 重新部署 Vercel

---

## 2. Supabase 建表

建议在 Supabase SQL Editor 中执行：

```sql
create table if not exists public.app_state (
  state_key text primary key,
  payload jsonb not null,
  updated_at timestamptz not null default now()
);
```

当前代码默认使用：

- 表名：`app_state`
- 主键标识：`global`

也就是说，整站状态会作为一条 JSON 记录保存在：

- `state_key = 'global'`

---

## 3. Row Level Security

如果你只是内部使用、先求跑通，可以先不开复杂权限。

但更推荐的做法是：

1. 打开 RLS
2. 只允许匿名读取 / 写入这一张状态表

示例：

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

注意：

这套策略适合你当前“共享后台面板”场景，但它本质上是匿名可写。  
如果以后要面向更广泛的人群，建议加登录权限控制。

---

## 4. Vercel 环境变量

在 Vercel 项目里新增以下环境变量：

```bash
VITE_SUPABASE_URL=你的 Supabase Project URL
VITE_SUPABASE_ANON_KEY=你的 Supabase anon key
VITE_SUPABASE_STATE_TABLE=app_state
VITE_SUPABASE_STATE_KEY=global
```

其中：

- `VITE_SUPABASE_URL` 和 `VITE_SUPABASE_ANON_KEY` 必填
- 后两个可不填，默认就是：
  - `app_state`
  - `global`

---

## 5. 本地调试

如果你想先本地试共享版，可以在项目根目录新建：

- `.env.local`

内容示例：

```bash
VITE_SUPABASE_URL=https://xxxx.supabase.co
VITE_SUPABASE_ANON_KEY=xxxxxxxx
VITE_SUPABASE_STATE_TABLE=app_state
VITE_SUPABASE_STATE_KEY=global
```

然后运行：

```bash
npm install
npm run dev
```

---

## 6. 当前共享实现方式

当前项目采用的是：

- **整站状态 JSON 统一存储**

也就是把当前网站中的主要状态整体保存到 Supabase：

- keepaRows
- yesterdayKeepaRows
- mappingRows
- monitorRows
- onlineRows
- sourceReports
- todayBuyBox
- yesterdayBuyBox
- keepaUploadReports

优点：

- 改造快
- 对现有页面侵入小
- 最适合当前项目从“本地缓存版”快速升级为“共享版”

缺点：

- 还不是细粒度表结构
- 后续如果要做多人并发编辑、变更审计、按模块历史回滚，需要继续拆表

---

## 7. 后续推荐升级方向

如果后面你们要长期多人协作，建议第二阶段再逐步拆成多张表：

1. mapping_rows
2. monitor_rows
3. online_monitor_rows
4. keepa_today_rows
5. keepa_yesterday_rows
6. app_meta / upload_reports

这样会更适合：

- 历史追溯
- 多人并发
- 精细权限
- 局部更新

但现阶段，先用 `app_state` 共享整站状态是最省事也最实用的方案。

---

## 8. 当前还没做完的部分

虽然共享骨架已经接入代码，但当前还没有完成以下事项：

1. 还没把 Supabase 环境变量真正配置到你的 Vercel
2. 还没在 Supabase 中实际建表
3. 还没验证“你上传一次，别人浏览器立即读到同一份数据”的线上闭环

所以现在项目处于：

- **代码已具备共享接入基础**
- **但还不算共享版最终完成**

---

## 9. 推荐下一步执行顺序

建议按这个顺序做：

1. 建 Supabase 项目
2. 执行建表 SQL
3. 配置 Vercel 环境变量
4. 重新部署 Vercel
5. 上传一份 Keepa / 映射 / 监控清单测试
6. 换浏览器或换电脑打开同一个 Vercel 链接验证共享

---

## 10. 给后续 Codex 的一句话

如果新的 Codex 对话要继续做共享版，可以直接告诉它：

> 当前项目已经接入 Supabase 共享状态骨架，代码在 `src/lib/sharedState.ts` 和 `src/App.tsx`，请继续完成 Supabase 共享部署闭环，并确保 Vercel 多用户访问看到的是同一份数据。

