import datetime
import webbrowser
import http.server
import socketserver
import requests
import json
import os
#非常重要提示！！！
#需要修改的地方app_id,app_secret,PORT
#端口和在申请应用程序时的回调地址一致，例如我的是http://localhost:7777/，则端口为7777
#第一次运行时在本文件所在目录创建名为time.txt和refreshtoken.txt的两个空文件存储变量。
#本程序仅实现了./me的个人信息获取，想要实现更多api，可以在get_message()函数中写。


#实现判断refresh_token过期时间，默认90天
#设置refresh_token过期时间
refresh_days=80
#记得修改应用信息
app_id=r''
app_secret=r''
redirect_uri=''
#记得修改端口
PORT = 53682

#以下内容无需修改！！！
#运行变量
path=""

def pdtoken():
    token=dqtoken()
    if (token==""):
        print("token为空，执行第一次运行")
        xrsj()
        code=get_code()
        access_token=get_token(code)
        get_message(access_token["access_token"])
        print("获取信息成功！")
        xrtoken(access_token["refresh_token"])
    else:
        print("token存在，开始判断是否过期")
        get_time()

def get_time():
    time1=datetime.datetime.now()
    time2=dqsj()
    cha=time2-time1
    xrsj()
    if(cha.days>=refresh_days):
        print("token已经接近过期时效，更新中")
        rstoken=dqtoken()
        shuaxintoken=refresh_access_token(rstoken)
        get_message(shuaxintoken["access_token"])
        print("获取信息成功！")
        xrtoken(shuaxintoken["refresh_token"])

    else:
        print("token在有效期内，正常运行")
        yxtoken=dqtoken()
        shuaxintoken=refresh_access_token(yxtoken)
        get_message(shuaxintoken["access_token"])
        print("获取信息成功！")
                

#重写http请求类
class SETHandler(http.server.SimpleHTTPRequestHandler):
      def do_GET(self):
         global path
         path=self.path

def get_code():
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={
          'client_id':app_id,
          'redirect_uri': redirect_uri
         }
    url='https://login.microsoftonline.com/common/oauth2/v2.0/authorize?'+"client_id="+data["client_id"]+"&response_type=code"+"&redirect_uri="+data["redirect_uri"]+"&scope=offline_access%20user.read%20mail.read"+"&response_mode=query"
    webbrowser.open(url)
    Handler = SETHandler
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        print("serving at port", PORT)
        httpd.handle_request()
    code=(path.split("code=")[1]).split("&session_state=")[0]
    return code

#获取access_token和refresh_token
def get_token(code):
    headers={'Content-Type':'application/x-www-form-urlencoded',
            }
    data={
          'code': code,
          'client_id':app_id,
          'client_secret':app_secret,
          'scope':'user.read',
          'grant_type':'authorization_code',
          'redirect_uri':redirect_uri
         }
    html = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    shuju=dict()
    shuju["refresh_token"]=jsontxt["refresh_token"]
    shuju['access_token']=jsontxt["access_token"]
    return shuju

#使用access_token访问资源
def get_message(access_token):
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        str1=requests.get(r'https://graph.microsoft.com/v1.0/me',headers=headers)
        dict1=json.loads(str1.text)
        print(dict1)
    except Exception as e:
        print(e)
#使用refresh_token刷新access_token
def refresh_access_token(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded',
            }
    data={
          'client_id':app_id,
          'client_secret':app_secret,
          'scope':'user.read',
          'grant_type':'refresh_token',
          'refresh_token':"",
          'redirect_uri':redirect_uri,
          'refresh_token':refresh_token
         }
    html = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    shuju=dict()
    shuju["refresh_token"]=jsontxt["refresh_token"]
    shuju['access_token']=jsontxt["access_token"]
    return shuju

def dqsj():
    with open(r"time.txt","r",encoding="UTF-8")as op1:
        gt=op1.read()
    dqtime=datetime.datetime.strptime(gt,'%Y-%m-%d')
    return dqtime

def xrsj():
    xrtime=datetime.datetime.now().strftime('%Y-%m-%d')
    with open(r"time.txt","w+",encoding="UTF-8")as op2:
        op2.write(xrtime)

def dqtoken():
    with open(r"refreshtoken.txt","r",encoding="UTF-8")as op3:
        token=op3.read()
    return token

def xrtoken(refresh_token):
    with open(r"refreshtoken.txt","w+",encoding="UTF-8") as op4:
        op4.write(refresh_token)

pdtoken()
os.system("pause")