# HyperFormula 环境配置（Windows + PowerShell）

本文档用于在本仓库中快速完成本地开发环境配置。

## 1. Node.js 版本

- 推荐使用 **Node.js 22 LTS**（`docs/guide/building.md` 中说明构建链路使用 Node 22 LTS）。
- 仓库中的 `.nvmrc` 当前是 `v16`，可视为历史残留；本仓库开发优先按 Node 22 LTS 配置。

## 2. 进入仓库并安装依赖

先进入**仓库根目录**（即本项目 `README.md` 所在目录），然后执行：

```powershell
npm.cmd ci
```

说明：

- 在 PowerShell 中如果直接 `npm` 遇到执行策略限制，优先使用 `npm.cmd`。
- 或者放开当前用户执行策略：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## 3. 验证基础开发环境

在仓库根目录执行：

```powershell
node -v
npm.cmd -v
npm.cmd run compile
npm.cmd run test:jest -- --runInBand smoke
```

判断标准：

- `compile` 通过：TypeScript 构建链路正常。
- `smoke` 通过：最小测试链路正常。

## 4. 常用开发命令

```powershell
npm.cmd run bundle-all      # 完整打包
npm.cmd run test            # lint + jest + browser tests
npm.cmd run docs:dev        # 本地文档开发服务器
```

## 5. 补充说明

- `test/README.md` 说明：本仓库只包含 smoke 测试，完整测试集不在此仓库内。
- `docs:code-examples:*` 脚本依赖 `bash`，如需运行请使用 Git Bash 或 WSL。
