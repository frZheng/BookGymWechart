# BookGymWechart
使用微信发送消息预约羽毛球场地

定时每周一8:00发送信息给管理员,预约周三的场地.
实现原理是打开微信后, 输入发送信息对象的名称后, 将消息粘贴到粘贴板, 按下altS发送消息.
对没打开微信的, 微信不在最顶层进行一些条件判断. 每天运行时因为太久不发送消息, 微信会自动退出.