# SendSalarybyPDF
Send salary report in PDF format via email

用最简单的脚本实现 通过邮件给员工发送PDF格式的工资单

脚本 １是用来生成ＰＤＦ工资单

脚本２ 是用来改善ＰＤＦ工资单

步骤：

１. 从表格工资单获取数据，参考模板，可增加列，一定要保留前的ＩＤ，这个ＩＤ可以决定在ＰＤＦ报告中显示的顺序

２. 按员工ＩＤ，生成对应ＴＸＴ 文本

3. 转换成ＰＤＦ报告

4. 遍历报告，通过邮件发送给对应的员工
