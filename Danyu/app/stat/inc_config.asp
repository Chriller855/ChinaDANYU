<%
'基本信息
countname	= "网站统计"	' 统计器名称
yesleft		= 1				' 是否显示左侧的导航，1表示是，0表示否
yestop		= 0				' 横导航栏放置位置，0表示不放，1表示放页顶，2表示放页底
yesto		= 0				' 是否在版权栏中显示“授权×××”使用字样，1表示显示，0表示不显示
connpath	= "count.mdb"	' 数据库存放位置，这里填写相对index.asp的相对路径
ipconnpath	= "ip.mdb"		' IP库存放位置。
bakconnpath	= "stats.mdb"	' 后备库存放位置。
whatcan		= 4				' 访客的权限等级，有关等级划分请看 readme 文档
CookieExpires=100			' cookie设置时间(天),默认为100天
addtime		= 0				' 服务器时间差，服务器时间和本地时间查多少？单位小时
							' 比如服务器时间是17:00，本地时间19:00，就设置2
							' 服务器时间17:00，本地时间15:00，就设置-2
old_count	= 0				' 在使用本统计系统前的访问量

							' 注意，防刷新和在线统计都将降低统计器效率

is_ipcheck	= true			' 是否启用防刷新功能，true表示启用，false表示不启用
is_online	= false			' 是否统计在线人数，true表示启用，false表示不启用

onlythesite1= ""			' 限制嵌入代码位置，保持空白表示不做限制
onlythesite2= ""			' 限制嵌入代码位置，保持空白表示不做限制
'此项不为空时，系统在执行时会检查被访问页面的URL是否以上述文字开头，
'不是以上述文字开头的网址一律不计数。
'比如 onlythesite1="http://www.01demo.com" ，则如有人在非http://www.01demo.com网页放置
'嵌入代码，则统计器并不计数。


'外观参数
FlashWidth	= 130			' FLASH图标宽度
FlashHeight	= 58			' FLASH图标高度
mPageSize	= 15			' 每页记录数（仅对需分页的页面）
mPrecision	= 5				' 精确到小数点后多少位


'站点信息
mURL		= "http://www.01demo.com"	' 站点连接
mName		= "01demo"				' 站点名称
mNameEn		= "01demo.com"			' 站点英文名称
masterName	= "01demo"					' 管理员名
masterpsw	= "admin"					' 管理员密码
masterEmail	= "zhangyu@01demo.com"			' 管理员电子邮件
SiteBrief	= ":)"	' 站点介绍，请勿使用英文的双引号和换行

'conn
set conn=server.createobject("adodb.connection")
DBPath = Server.MapPath(connpath)
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath


%>