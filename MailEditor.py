# Mail Editor

import os
import time
import pyperclip
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import shutil
from configparser import ConfigParser

def code():
    ct=time.localtime()
    code=str(hex(int(time.strftime('%Y%m%d%H%M%S',ct))))
    return code
cd=code()

print('读取配置...')
conf = ConfigParser()
conf.read('mail.ini')
mail_host=conf.get("info","mail_host")
mail_user=conf.get("info","mail_user")
mail_pass=conf.get("info","mail_pass")
sender=conf.get("info","sender")
receivers=conf.get("info","receivers")
print('配置读取完成。')
input('\n即将发送邮件，请确认用户名：'+receivers+'。\n如不符合，请修改软件目录下mail.ini 文件。')

names = os.listdir('inbox')
s=len(names)
if s==0:
    input('\n目标文件夹为空。即将退出。')
    exit()
else:
    print('附件已读取。共',s,'个文件。\n')

list0=[]
order=0
for i in names:
    order+=1
    while True:
        a=int(input('(第 '+str(order)+' 个，共 '+str(s)+' 个)为 '+i+' 文件选择打印选项。单面打印请填1，双面打印请填2，仅保存选3。'))
        if a==1:
            list0.append('单面打印')
            break
        elif a==2:
            list0.append('双面打印')
            break
        elif a==3:
            list0.append('仅保存')
            break
        else:
            print('警告：非法字符。')

timestr=str(input('\n截止时间：'))
t=time.strptime(timestr, '%Y%m%d%H%M%S')

text=''
log=''
log+='这是MailEditor在'+time.strftime('%Y'+'年'+'%m'+'月'+'%d'+'日'+'%H'+'时'+'%M'+'分'+'%S'+'秒', time.localtime())+'的日志，请求编号为'+cd+'。\n'
text+="""<div>
    <includetail>
        <table style="font-family: Segoe UI, SegoeUIWF, Arial, sans-serif; font-size: 12px; color: #333333; border-spacing: 0px; border-collapse: collapse; padding: 0px; width: 580px; direction: ltr">
            <tbody>
            <tr>
                <td style="font-size: 10px; padding: 0px 0px 7px 0px; text-align: right">
                    Mail Editor v1.0.1
                </td>
            </tr>
            <tr style="background-color: #0078D4">
                <td style="padding: 0px">
                    <table style="font-family: Segoe UI, SegoeUIWF, Arial, sans-serif; border-spacing: 0px; border-collapse: collapse; width: 100%">
                        <tbody>
                        <tr>
                        <tr>
                            <td style="font-size: 20px; color: #FFFFFF; padding: 0px 22px 18px 22px" colspan="3">
                                Mail Editor v1.0.1
                            </td>
                        </tr>
                            <td style="padding: 0px; width: 100%">
                            </td>
                        </tr>
                        <tr>
                            <td style="font-size: 38px; color: #FFFFFF; padding: 12px 22px 4px 22px" colspan="3">
                                打印请求
                            </td>
                        </tr>
                        <tr>
                            <td style="font-size: 20px; color: #FFFFFF; padding: 0px 22px 18px 22px" colspan="3">
                                打印请求需要您处理
                            </td>
                        </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="padding: 30px 20px; border-bottom-style: solid; border-bottom-color: #0078D4; border-bottom-width: 4px">
                    <table style="font-family: Segoe UI, SegoeUIWF, Arial, sans-serif; font-size: 12px; color: #333333; border-spacing: 0px; border-collapse: collapse; width: 100%">
                        <tbody>
                        <tr>
                            <td style="font-size: 12px; padding: 0px 0px 5px 0px">
                                请根据需求完成该打印请求
                                <ul style="font-size: 14px">"""
for i in range (0,s):
    if i==0:
        text+='\n                                    <li style="padding-top: 10px">\n                                        第'+str(i+1)+'份：'+str(list0[i])+'，'+str(names[i])+'。\n                                    </li>'
    else:
        text+='\n                                    <li>\n                                        第'+str(i+1)+'份：'+str(list0[i])+'，'+str(names[i])+'。\n                                    </li>'

text+=""" \n                               </ul>
                            </td>
                        </tr>
                        <tr>
                            <td style="font-size: 12px; padding: 0px 0px 15px 0px">"""
text+=' \n                               共'+str(s)+'份文件。请在'+time.strftime('%Y'+'年'+'%m'+'月'+'%d'+'日'+'%H'+'时'+'%M'+'分'+'%S'+'秒', t)+'前打印完成。'
text+="""\n                            </td>
                        </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="padding: 35px 0px; color: #B2B2B2; font-size: 12px">"""
text+='\n                    请求编号：'+cd
text+='\n                    <br>'
text+='\n                    请求时间：'+time.strftime('%Y'+'年'+'%m'+'月'+'%d'+'日'+'%H'+'时'+'%M'+'分'+'%S'+'秒', time.localtime())
text+="""\n                    <br>
                    Tresol
                    <br>
                    Mail Editor v1.0.1
                </td>
            </tr>
            <tr>
                <td style="padding: 0px 0px 10px 0px; color: #B2B2B2; font-size: 12px">
                    版权所有 Tresol
                    <br>
                    Mail Editor v1.0.1 及其附属文件已在 <a href="http://github.com/Tresol/MailEditor" style="color: #0044CC">Github</a> 上开放源代码，遵守 <a href="https://www.gnu.org/licenses/gpl-3.0.html#license-text" style="color: #0044CC">GNU通用公共许可证</a> 协议。
                    <br>
                    <a href="http://github.com/Tresol" style="color: #0044CC">需要帮助? 请与Mail Editor开发者联系</a>
                </td>
            </tr>
            </tbody>
        </table>
    </includetail>
</div>"""

log+='文档HTML草稿为：\n\n'+text+'\n'

message = MIMEMultipart()
message.attach(MIMEText(text,'html','utf-8'))       
message['Subject'] = '打印请求'+cd 
message['From'] = sender  
message['To'] = receivers

mailver=int(input('\n选0发送邮件至指定地址。'))
if mailver==0:
    log+='用户发出了发送邮件的请求。\n'
    for i in range (0,s):
        att1=MIMEText(open('inbox\\'+names[i], 'rb').read(), 'base64', 'gb2312')
        att1["Content-Type"] = 'application/octet-stream'
        att1.add_header("Content-Disposition", "attachment", filename=("gbk", "",names[i]))
        message.attach(att1)
    input('\n即将发送邮件，请确认用户名：'+receivers+'。\n')
    log+='邮件发送账号：'+sender+'；邮件收件人：'+receivers +'。\n'
    try:
        smtpObj = smtplib.SMTP() 
        smtpObj.connect(mail_host,25)
        smtpObj.login(mail_user,mail_pass) 
        smtpObj.sendmail(
            sender,receivers,message.as_string()) 
        smtpObj.quit() 
        print('\n邮件发送成功。')
        log+='邮件发送成功。\n'
    except smtplib.SMTPException as e:
        print('\n邮件发送失败。错误代码：',e)
        log+='邮件发送失败。错误代码：'+e+'\n'

print('\n正在进行存档……')
for i in names:
    if os.path.exists('archives\\'+cd)==False:
        os.makedirs('archives\\'+cd)
    shutil.move('inbox\\'+i,'archives\\'+cd)
log+='已存档至'+'archives\\'+cd+'。\n'
print('\n存档完成。\n创建日志...')

file = open('archives\\'+cd+'\\'+cd+'.html', "w", encoding='utf-8')
file.write(text)
file.close()

file = open('archives\\'+cd+'\\'+cd+'.txt', "w", encoding='utf-8')
file.write(log)
file.close()
print('日志创建完成。')
input()
