使用说明：
	1.如果是初始下载,请检查是否安装node；请检查是否有相应的依赖文件node_modules，如果没有运行npm i；
	2.savetime.cmd配置脚本文件的执行路径（重要）；
	3.设置系统关机自动执行savetime.cmd脚本：
		在 开始菜单 运行 gpedit.msc 打开组策略编辑器。 
		到左边找到 计算机配置-Windows配置-启动/关机脚本 双击右边的关机脚本 
		在弹出的对话框中选择添加 找到savetime.cmd文件，确定；
	4.关机时Excel表格默认保存在e盘根目录下，可自己配置savetime.cmd修改；