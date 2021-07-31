import itchat,time,sys,logging

#item['NickName','RemarkName','Sex','Province','City','PYInitial','PYQuanPin','UserName','Signature','HeadImgUrl'
#'DisplayName',ChatRoomId
#
# def loginComplete():
#     logger.info('Login completed')
# def logoutComplete():
#     logger.info('Logout completed')




class WechatAgent:
    def __init__(self,logger):
        self.logger = logger
        itchat.auto_login(hotReload=True) # loginCallback=loginComplete, exitCallback=logoutComplete)


        pass

    ## register text_message_handler function as the handler whenever receiving a simple text message
    @itchat.msg_register([itchat.content.TEXT],isFriendChat=True)
    def text_message_handler(self, msg):
        logger.info(msg['Text'] + ' received from user:' + msg['FromUserName'])
        return '收到消息:“' + msg['Text'] + '” 抱歉，机主目前没空，这是系统自动回复'          # reply any text message

    @itchat.msg_register(itchat.content.PICTURE,isFriendChat=True)
    def picture_message_handler(self, msg):
        logger.info(msg)
        #msg.FromUserName,FileName,MsgType==3,'NickName','RemarkName'

    @itchat.msg_register(itchat.content.VIDEO,isFriendChat=True, isGroupChat=False)
    def video_message_handler(self, msg):
        logger.info(msg)
        return '来自' + msg['FromUserName'] +'的视频收到，自动应答'

    def quit(self):
        itchat.logout()

    def send_msg(self, nickName , mesg):
        try:
            friends = itchat.search_friends(nickName=nickName)

            itchat.send(mesg, friends[0]['UserName'])    # send a message to userName
            return 0
        except Exception as e:
            self.logger.error(e)


if __name__=='__main__':
    logger = logging.getLogger()
    stream = logging.StreamHandler()
    # formatter = logging.Formatter()
    formatter = logging.Formatter('%(asctime)s: %(levelname)s %(message)s')
    stream.setFormatter(formatter)
    logger.addHandler(stream)
    logger.setLevel(logging.DEBUG)

    wx = WechatAgent(logger)
    wx.send_msg(nickName='随风往事',mesg='hello,this is a test message,收到请回复')
    wx.quit()

    # friends = itchat.get_friends(update=True)[0:]
    # my_friend = itchat.search_friends(nickName='随风往事')

    # male = female = other = 0
    # for i in friends[1:]:
    #     sex = i['Sex']
    #     if sex == 1:
    #         male += 1
    #     elif sex == 2:
    #         female += 1
    #     else:
    #         other += 1
    # total = len(friends[1:])
    # hxm = friends[0]['UserName']
    # my_chatroom = itchat.search_chatrooms(name='天涯海北中二')
    # logger.info(my_chatroom)

    # result = 'you have {} friends, {} male, {} female,{} unspecified'.format(total,male,female,other)
    # logger.info(result)
    # filename = 'bell.wav'
    # itchat.send('sent from myself:'+result, myself) #'filehelper')   # send a message to self filehelper
    # # itchat.send('@%s@%s' % ('fil', filename), 'filehelper')     # send a file by the name of filename to self
    # itchat.run(debug=True)


