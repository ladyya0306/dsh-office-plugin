# dsh-office-workflows

“办公程序复用与精确填报”是一个面向 DSH 的办公插件包。它提供办公文件读取、程序保存与复用、独立产物检查等能力；默认关闭普通程序执行和学习记录，实际能力仍受 DSH 配置、工作区和宿主检查约束。

当前冻结包版本为 `0.3.0-rc.7`，已验收DSH运行版本为 `0.2.0-rc.2`。源码、界面资源和独立离线插件ZIP由本仓库发布；具体模型和目标电脑按各自任务验收。

## 安装

### CLI 或本地目录

联网电脑使用DSH官方插件管理入口添加GitHub来源，或在命令行执行：

```powershell
dsh plugin --profile web add github:ladyya0306/dsh-office-plugin#v0.3.0-rc.7
```

离线电脑从[Releases](https://github.com/ladyya0306/dsh-office-plugin/releases)取得插件ZIP，解压到新目录，使用预置依赖的DSH离线包执行“离线安装办公插件.cmd”，输入插件绝对目录；或执行：

```powershell
dsh plugin --profile web add --offline --ignore-scripts --config.auto-install-peers=false file:D:/DSH插件/dsh-office-workflows-0.3.0-rc.7
```

安装后按DSH要求重启，核对名称、rc7版本、包与组件开关及9项中文设置。插件ZIP提供本项目程序和资源，DSH、Node、Python和包管理器由DSH运行包提供。安装目录保留供加载。运行时无需访问GitHub。

停用及卸载使用DSH插件管理器；随离线DSH也提供“离线卸载办公插件.cmd”。来源、模板、产物、方法库和运行记录保存在独立目录，卸载后继续保留。

## 运行边界

- 默认程序执行关闭：`nativeScriptMode=disabled`、`learningEnabled=false`、`captureOrdinaryPython=false`。
- 个人兼容模式 `personal-compatible` 只适用于获准的个人电脑测试，并应记录实际 Windows 约束。
- 行内受控模式必须在具备相应宿主控制的环境中单独验证；本目录不自行放宽工作区、工具白名单、网络、系统命令或审计边界。
- 输出写入新文件，保护来源文件和模板原件。输出生成后仍需独立重新打开和检查。
- 修改处理程序或参数化后，重新执行、核对实际产物，再保存新版本；下一会话按本次位置复用。

## 隐私与许可

公开仓库仅保存插件源码、资源和公开说明；用户资料、方法库、会话、凭据、实际配置及运行记录留在本机。

项目持有人授权此次上传，原创代码当前保留权利，开源使用与再分发许可另行明确。第三方代码和依赖遵循原许可，来源说明见[THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

## 当前版本与验收

业务程序来自冻结产品包rc7，66个源码、worker、适配模块及界面配置文件与原冻结摘要逐项一致。package.json添加公开仓库信息。

此次Windows离线交付的官方插件管理器完成安装、启用、停用及卸载六步检查，包管理器下载数量为0；随包Python完成真实XLSX写入、前导零和真实0、公式与格式保留、覆盖拒绝、路径越界拒绝、错误值拒绝、非法操作说明和原件校验等8项检查。具体行内模型地址、协议、证书、权限和陌生资料由目标电脑继续验收。

普通会话减少重编程、保存前空白格样式检查及单Excel方法元数据进一步分离列为后续研发内容。
