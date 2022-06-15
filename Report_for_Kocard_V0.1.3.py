"""
Reports Tool for Kocard
Version: 0.1.3
Author: ZessO
"""
import copy
import re
import time
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_RIGHT
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib import colors
import tkinter
from tkinter import filedialog

#TTF字体注册
pdfmetrics.registerFont(TTFont('Menlo', 'Menlo-Pingfang.ttf'))

#GUI
root = tkinter.Tk()

root.title('Reports Tool for Kocard by ZessO')
root.geometry('435x180')
root.resizable(width=False, height=False)

labelTitle = tkinter.Label(root, text='请输入项目名称：', font=('', 11))
labelTitle.place(x=0, y=0)

titleEntry = tkinter.Entry(root, show='', font=('', 11))
titleEntry.place(x=0, y=20, width=250, height=25)

labelName = tkinter.Label(root, text='请输入已校验的备份数量：', font=('', 11))
labelName.place(x=0, y=60)

nameEntry = tkinter.Entry(root, show='', font=('', 11))
nameEntry.place(x=0, y=80, width=250, height=25)

labelPath = tkinter.Label(root, text='请选择MHL文件', font=('', 11))
labelPath.place(x=0, y=120)

entry_text = tkinter.StringVar()
entry = tkinter.Entry(root, textvariable=entry_text, font=('', 10), width=30, state='readonly')
entry.place(x=0, y=140, width=250, height=25)

def getPath():
    path = filedialog.askopenfilenames(title='请选择MHL文件')
    entry_text.set(path)

def getTextInput(textName):
    result = textName.get()    #获取文本输入框的内容
    return result

def genReport():
    allPath = getTextInput(entry_text)
    allPath = allPath.strip('()')
    pathList = allPath.split()
    path_num = len(pathList)
    for i in range(0, path_num):
        pathList[i] = pathList[i].strip(',')
        pathList[i] = pathList[i].strip("'")
    # print(projectPath)

    projectTitle = getTextInput(titleEntry)
    backups = getTextInput(nameEntry)

    if path_num == 0:
        tkinter.messagebox.showinfo(title='提示', message='未检测到路径！')
        exit()
    else:
        for i in range(path_num):
            projectPath = pathList[i]

            if projectTitle == '' or backups == '' or projectPath == '':
                tkinter.messagebox.showinfo(title='提示', message='请输入完整信息！')
                exit()

            # 单项信息获取函数
            def single_get(what_type):
                single = re.findall(what_type, log_content)
                single[0] = single[0][single[0].find(">") + 1:single[0].rfind("<")]
                single_out = single[0]
                return single_out

            # 文件获取及log读取
            log_place = projectPath

            log_source = open(log_place, encoding='utf-8')
            log_content = log_source.read()
            log_source.close()
            # print(log_content)

            # 基础判断
            status = 'Success'

            # 判断创建MHL的操作系统
            system_mhl = False
            if re.search(r'\\', log_content) != None:
                system_mhl = True
            # print(system_mhl)

            # 获取项目信息
            project_name = projectTitle
            verified = backups

            # 获取当前时间
            local_time = time.strftime('%Y/%m/%d %H:%M', time.localtime(time.time()))

            # 单项信息获取
            # 卷名获取
            source_name_get = log_place[log_place.rfind('/'):]
            source_name = source_name_get[1:source_name_get.find('_')]
            # print(source_name)

            # 开始时间获取
            start_time = single_get(r'<startdate>.*<')
            print(type(start_time))
            start_time = start_time.replace('T', '\u3000')
            start_time = start_time.replace('Z', '')
            start_time = start_time.replace('-', '/')
            # print(start_time)

            # 结束时间获取
            finish_time = single_get(r'<finishdate>.*<')
            finish_time = finish_time.replace('T', '\u3000')
            finish_time = finish_time.replace('Z', '')
            finish_time = finish_time.replace('-', '/')
            # print(finish_time)

            # Hash类型获取
            hash_type = 'xxhash'

            # 数据转换
            def hum_convert(value):
                units = ["B", "KB", "MB", "GB", "TB", "PB"]
                size = 1024.0
                for i in range(len(units)):
                    if (value / size) < 1:
                        return "%.2f%s" % (value, units[i])
                    value = value / size

            # 正则批量取数据函数
            def batch_get(which_part):
                batch = re.findall(which_part, log_content)
                # print(batch)
                for i in range(len(batch)):
                    batch[i] = batch[i][batch[i].find(">") + 1:batch[i].rfind('<')]
                return batch

            # 取单个文件大小
            single_size = batch_get(r'<size>.*<')
            # print(single_size)
            single_size_original = single_size * 1
            # print(single_size_original)
            total_size = [0]
            #print(len(single_size_original))
            for i in range(0, len(single_size_original) - 1):
                total_size[0] += int(single_size_original[i])

            # 涉数据转换部分处理
            total_size_out = total_size[0]
            total_size_out = hum_convert(float(total_size_out))
            # print(total_size_out)

            for i in range(0, len(single_size)):
                single_size[i] = hum_convert(int(single_size[i]))
            # print(single_size)

            # 取文件名
            clip_name = batch_get(r'<file>.*<')
            print(clip_name)
            simple_name = copy.deepcopy(clip_name)
            if system_mhl == True:
                for i in range(len(simple_name)):
                    simple_name[i] = simple_name[i][simple_name[i].rfind('\\') + 1:]
            else:
                for i in range(len(simple_name)):
                    simple_name[i] = simple_name[i][simple_name[i].rfind('/') + 1:]

            for i in range(len(simple_name)):
                simple_name[i] = simple_name[i][:simple_name[i].rfind(".")]

            # 总文件数获取
            total_files_transferred = len(clip_name)
            # print(total_files_transferred)

            # Hash值
            hash_values = batch_get(r'<xxhash>.*<')

            # Clips Overview模块
            # 获取文件名的后缀名
            def get_suffix(filename):
                # 从字符串中逆向查找.出现的位置
                pos = filename.rfind('.')
                # 通过切片操作从文件名中取出后缀名
                return filename[pos + 1:] if pos > 0 else ''

            video_list = ['MOV', 'MXF', 'MP4', 'ARI', 'ARX', 'MPG', 'MPEG', 'DNG', 'BRAW', 'CRM', 'MTS', 'VRW', 'R3D',
                          'CINE',
                          'AVI', 'MKV', 'RMF', 'KRW', 'RED']
            audio_list = ['WAV', 'MP3', 'AAC', 'M4A', 'APE', 'FLAC', 'WMA']
            files_count = [0, 0, 0]
            size_count = [0, 0, 0]
            for i in range(len(clip_name)):
                if (get_suffix(clip_name[i]).upper() in video_list) == True:
                    files_count[0] += 1
                    size_count[0] += int(single_size_original[i])
                    # print(files_count)
                elif (get_suffix(clip_name[i]).upper() in audio_list) == True:
                    files_count[1] += 1
                    size_count[1] += int(single_size_original[i])
                else:
                    files_count[2] += 1
                    size_count[2] += int(single_size_original[i])
            print(size_count)

            for i in range(0, 3):
                size_count[i] = hum_convert(size_count[i])
            print(size_count)

            def formReport():
                # PDF文件创建
                pdf = SimpleDocTemplate(source_name + '_' + 'DMT Report' + '.pdf',
                                        pagesize=landscape(A4),
                                        title=source_name + '_' + 'DMT Report',
                                        rightMargin=45,
                                        leftMargin=45,
                                        topMargin=25,
                                        bottomMargin=25)

                story = []

                # 定义段落Style
                styleARight = ParagraphStyle(
                    name='Normal',
                    fontSize=12,
                    alignment=TA_RIGHT,
                )

                styleA = ParagraphStyle(
                    name='Normal',
                    fontSize=12,
                    leftIndent=-15,
                )

                styleB = ParagraphStyle(
                    name='Normal',
                    fontSize=28,
                    fontName='Menlo',
                    leftIndent=15,
                )

                styleC = ParagraphStyle(
                    name='Normal',
                    fontSize=10,
                    leftIndent=15,
                    fontName='Menlo',
                    textColor=colors.Color(red=(102.0 / 255), green=(102.0 / 255), blue=(102.0 / 255)),
                )

                styleD = ParagraphStyle(
                    name='Normal',
                    fontSize=10,
                    leftIndent=97,
                    fontName='Menlo',
                    textColor=colors.Color(red=(102.0 / 255), green=(102.0 / 255), blue=(102.0 / 255)),
                )

                styleE = ParagraphStyle(
                    name='Normal',
                    fontSize=10,
                    leftIndent=-5,
                    fontName='Menlo',
                )

                # 添加Flowables
                para = [
                    [Paragraph('Clips Report', style=styleA), Paragraph(local_time, style=styleARight)]
                ]
                paraTbl = Table(para)
                story.append(paraTbl)
                story.append(Spacer(1, 36))

                para = Paragraph(source_name, style=styleB)
                story.append(para)
                story.append(Spacer(1, 32))

                para = Paragraph(project_name, style=styleC)
                story.append(para)
                story.append(Spacer(1, 5))

                para = Paragraph('Offloaded between' + ' ' + start_time, style=styleC)
                story.append(para)
                story.append(Spacer(1, 5))

                para = Paragraph('and' + ' ' + finish_time, style=styleD)
                story.append(para)
                story.append(Spacer(1, 20))

                para = Paragraph('Clips Overview', style=styleE)
                story.append(para)
                story.append(Spacer(1, 10))

                data = [[' ', 'Clips', 'Files', 'Size'],
                        ['Video Clips', files_count[0], files_count[0], size_count[0]],
                        ['Audio Clips', files_count[1], files_count[1], size_count[1]],
                        ['Other Clips', '0', files_count[2], size_count[2]],
                        ['Total', files_count[0] + files_count[1], total_files_transferred, total_size_out]
                        ]

                tblstyle = TableStyle([('LINEBEFORE', (0, 0), (-1, -1), 3, colors.white),
                                       ('LINEAFTER', (0, 0), (-1, -1), 3, colors.white),
                                       ('FONTNAME', (0, 0), (-1, -1), 'Menlo'),
                                       ('FONTSIZE', (0, 0), (-1, 0), 11.5),
                                       ('FONTSIZE', (0, 1), (-1, -1), 10),
                                       ('VALIGN', (0, 0), (-1, 0), 'TOP'),
                                       ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                                       ('ROWBACKGROUNDS', (0, 0), (-1, -1),
                                        [colors.Color(red=(238.0 / 255), green=(238.0 / 255), blue=(238.0 / 255)),
                                         colors.white]),
                                       ])

                tbl = Table(data, rowHeights=(25, 20, 20, 20, 20), hAlign='LEFT')
                tbl.setStyle(tblstyle)
                story.append(tbl)
                story.append(Spacer(1, 20))

                tdata = []
                data = []
                data.append('Name')
                data.append('File Type')
                data.append('File Size')
                data.append('Hash Values')
                data.append('Verified Backups')
                data.append('Status')
                tdata.append(data)

                for i in range(0, int(total_files_transferred)):
                    data = []
                    data.append(simple_name[i])
                    # print(clip_name[i])
                    data.append(get_suffix(clip_name[i]))
                    # print(get_suffix(clip_name[i]))
                    data.append(single_size[i])
                    # print(single_size[i])
                    data.append(hash_type + ': ' + hash_values[i])
                    # print(hash_type + ': ' + hash_values[i])
                    data.append(verified)
                    data.append(status)
                    tdata.append(data)
                    print(tdata)

                tblstyle = TableStyle([('LINEBEFORE', (0, 0), (-1, -1), 3,
                                        colors.Color(red=(247.0 / 255), green=(247.0 / 255), blue=(247.0 / 255))),
                                       ('LINEAFTER', (0, 0), (-1, -1), 3,
                                        colors.Color(red=(247.0 / 255), green=(247.0 / 255), blue=(247.0 / 255))),
                                       ('FONTNAME', (0, 0), (-1, -1), 'Menlo'),
                                       ('FONTSIZE', (0, 0), (-1, -1), 10),
                                       ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                                       ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                                       ('ROWBACKGROUNDS', (0, 0), (-1, -1),
                                        [colors.Color(red=(238.0 / 255), green=(238.0 / 255), blue=(238.0 / 255)),
                                         colors.white]),
                                       ])

                tbl = Table(tdata, colWidths=(210, 85, 90, 180, 120, 70), hAlign='LEFT')
                tbl.setStyle(tblstyle)
                story.append(tbl)

                pdf.build(story)

            formReport()
        tkinter.messagebox.showinfo(title='提示', message='成功生成报告！')


button = tkinter.Button(root, text='选择路径', command=getPath)
button.place(x=260, y=140)

goOn = tkinter.Button(root, text='生成报告', command=genReport)
goOn.place(x=345, y=140)

root.mainloop()