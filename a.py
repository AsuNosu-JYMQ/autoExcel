import speech_recognition as sr
r = sr.Recognizer()
with sr.Microphone() as source:
    print("请开始说话...")
    # 监听麦克风输入
    audio = r.listen(source)
    try:
        # 使用Google Speech Recognition引擎进行语音识别
        text = r.recognize_google(audio, language="zh-CN")
        print("识别结果：" + text)
    except sr.UnknownValueError:
        print("无法识别语音")
    except sr.RequestError as e:
        print("请求出错：" + str(e))







from langchain_openai import ChatOpenAI
llm = ChatOpenAI(temperature=0.1,api_key='sk-a0e52a9c67b44d1d802754e539f5229b',model='deepseek-v3',base_url='https://dashscope.aliyuncs.com/compatible-mode/v1')
def ans_instruct(prompt):
    script =''
    for chunk in llm.stream(prompt):
        script += chunk.content
        print(chunk.content,flush=True,end='')
    return script


prompt_template='''
你是一个 Excel 助手。  
接下来我会提供两部分内容：  
1. **任务指示**
2. **表格结构与示例数据**（ 

请根据“任务指示”结合“表格结构与示例数据”生成详细的操作步骤。  
- 每一步都要用自然流畅的语言描述；  
- 涉及单元格、区域、函数、筛选、排序、复制等操作时，一定要标明精确的单元格或区域（如“A1”、“B2:D10”、“Sheet1!E1”）；  
- 步骤要编号，且保证按序可复现；  
- 如果需要创建新列或新工作表，也要说明名字和位置，例如“在第2列（B列）插入新列，命名为‘总计’”；  
- 对于复杂公式，要完整写出公式并标明输入位置；  
- 如果有依赖关系（例如先筛选再复制），请在每一步中体现。  

---  
'''
instruction =text
excel_info ='excel中没有任何内容'
prompt = f'{prompt_template}/n任务指示为：/n{instruction}/n表格结构与示例数据为：/n{excel_info}'
excel_instruction =  ans_instruct(prompt)



##second part
def ans_code(prompt):
    script =''
    for chunk in llm.stream(prompt):
        script += chunk.content
        print(chunk.content,flush=True,end='')
    return script
#  1. 在当前工作簿中新建一个名为“统计表”的工作表；
#    2. 从“原始数据”表里，把 A 列和 B 列所有非空行复制到“统计表”；
#    3. 在“统计表”末尾添加一行，计算 A 列的平均值和 B 列的总和；
prompt=f'''
你是一位精通 Excel VBA 的专家，生成的代码要能直接在 Excel 里运行，带有必要注释，并只输出代码块，不要多余说明。
请帮我写一个 VBA 宏，功能是：
    {excel_instruction }
 
	  
   要求：
   - 代码放在标准模块（Module1）里；
   - 给主要步骤加注释；
   - 最后只返回一个完整的 Sub 过程。"

输出示例：

Sub GenerateReport()
    ' 1. 新建“统计表”工作表
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("统计表")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "统计表"
    End If
    On Error GoTo 0

    ' 2. 复制“原始数据”表 A、B 列
    Dim src As Worksheet
    Set src = ThisWorkbook.Worksheets("原始数据")
    Dim lastRow As Long
    lastRow = src.Cells(src.Rows.Count, "A").End(xlUp).Row

    src.Range("A1:B" & lastRow).Copy Destination:=ws.Range("A1")

    ' 3. 在末尾添加统计行
    Dim tgtLast As Long
    tgtLast = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(tgtLast, "A").Value = "平均值："
    ws.Cells(tgtLast, "B").Formula = "=AVERAGE(A1:A" & tgtLast - 1 & ")"
    ws.Cells(tgtLast, "C").Value = "总和："
    ws.Cells(tgtLast, "D").Formula = "=SUM(B1:B" & tgtLast - 1 & ")"
End Sub



#最后注意：你的回答只包含代码。代码之外不要有任何内容。
'''
script =  ans_code(prompt)

with open(r"C:\Users\吉云拉觉\Desktop\vba.txt",'w',encoding='utf-8') as f:
    f.write(script)


#third part

import win32com.client as win32
import os
import time

# 1. 启动 Excel 应用
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True  # 必须设为 True，才能看到过程

# 2. 打开工作簿（xlsm 格式）
workbook_path = os.path.abspath(r"C:\Users\吉云拉觉\Desktop\新建 Microsoft Excel 工作表.xlsx")
wb = excel.Workbooks.Open(workbook_path)

# 3. 插入 VBA 代码（编码问题可切换为 encoding='gbk' 或用 try）
with open(r"C:\Users\吉云拉觉\Desktop\vba.txt", "r", encoding="utf-8") as f:
    vba_code= ''
    switch =0
    for line in f.readlines():
        if switch ==1:
            vba_code += line
            if 'End Sub' in line:
                switch = 0
        elif 'Sub' in line:
            vba_code += line
            macro_name = line[4:-3]
            print(macro_name)
            print(line)
            switch = 1
vb_module = wb.VBProject.VBComponents.Add(1)  # 1 = 模块
vb_module.CodeModule.AddFromString(vba_code)

# 4. 运行宏
excel.Application.Run(macro_name)  # 这里必须写对宏名称

# 5. 给一点时间让你看到执行过程（可选）
time.sleep(5)  # 观察5秒，避免太快关闭

# 6. 保存并关闭
wb.Save()
# wb.Close()  # 如想保留Excel窗口可先注释此行
# excel.Quit()  # 如果你想保留 Excel 窗口，也可注释这行





