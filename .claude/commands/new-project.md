用户要初始化一个新项目。

先询问：
1. 项目名称（英文小写，如 web3）
2. 项目描述（一句话）
3. 主要技术栈

然后执行：
1. 复制 _template/ 目录结构，创建以下文件：
   - <name>/context.md（填入项目名和描述）
   - <name>/bugs/INDEX.md
   - <name>/bugs/raw/（空目录，创建 .gitkeep）
   - <name>/requirements/raw/（空目录，创建 .gitkeep）
   - <name>/requirements/summaries/（空目录，创建 .gitkeep）

2. 更新 CLAUDE.md 的项目一览表，追加新项目一行

完成后提示用户：
- 若有专属设备，在 shared/devices/ 添加设备知识文件
- 将项目代码目录放入 <name>/ 下
