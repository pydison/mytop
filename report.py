#!/usr/bin/python
# -*- coding: utf-8 -*-
#import sys
#reload(sys)
#sys.setdefaultencoding('utf8')
import web
import os
import os.path
import MySQLdb

conn = MySQLdb.connect(host='localhost', port=3306, user='root', passwd='', db='itop', charset='utf8')
cur = conn.cursor()

class ATTA(object):
    """
    下载附件  
    """
    def GET(self,id,attaname):
        try:
            attaname=attaname.encode('utf-8')
            self.atta_name = ''.join(('/var/www/mytop/static/',str(id),'/',attaname))
            atta_= open(self.atta_name, "r+")
            web.header('Content-Type', 'application/octet-stream', 'charset = utf-8')
            web.header('Content-disposition', 'attachment; filename = %s' % attaname)
            content = atta_.read()
            return content
        except Exception as err:
            return err

class PDF(object):
    """
    生成pdf并下载
    """
    def __init__(self):
        """
        初始化
        """
        from reportlab.lib.pagesizes import letter
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle, TA_CENTER
        from reportlab.lib.units import inch
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, PageBreak
        import reportlab.rl_config
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        reportlab.rl_config.warnOnMissingFontGlyphs = 0
        pdfmetrics.registerFont(TTFont('song', '/var/www/mytop/static/simsun.ttc'))
        pdfmetrics.registerFont(TTFont('hei', '/var/www/mytop/static/simhei.ttf'))
        
        from reportlab.lib import fonts
        fonts.addMapping('song', 0, 0, 'song')
        fonts.addMapping('song', 0, 1, 'song')
        fonts.addMapping('song', 1, 0, 'hei')
        fonts.addMapping('song', 1, 1, 'hei')
        
        import copy
        stylesheet = getSampleStyleSheet()
        normalStyle = copy.deepcopy(stylesheet['Normal'])
        normalStyle.fontName = 'song'
        normalStyle.fontSize = 10

        self.width, self.height = letter
        self.width -= 50
        self.styles = getSampleStyleSheet()
        self.urgencys = {4:'低', 3:'中', 2:'高', 1:'紧急'}
        self.statusdict = {'assigned':'指派', 'resolved':'已解决', 'closed':'已关闭', 'escalated_ttr':'超时'}
        self.tdict = {'R':'iview_UserRequest', 'I':'iview_Incident', }

    def GET(self, ref):
        '''
        主程序入口
        '''
        id_ = int(ref[2:])
        table = self.tdict[ref[0]]
        sqltext = ''.join(("select id, ref, title, status, urgency, start_date, resolution_date, \
sla_ttr_passed, sla_ttr_over, time_spent, org_id_friendlyname, caller_id_friendlyname, \
agent_id_friendlyname, team_id_friendlyname, service_id_friendlyname, servicesubcategory_id_friendlyname, \
description, private_log, pending_reason, solution from ", table, " where id = %s;"))
        cur.execute(sqltext, (id_, ))
        self.tick = cur.fetchone()
        cur.execute('select contents_filename from iview_Attachment where item_id = %s', (id_, ))
        atts = [i[0] for i in cur.fetchall()]
        self.atts = '<br/>'.join(atts)

        cur.execute("select contact_id from ilnkcontacttoticket where ticket_id = %s", (id_, ))
        contacts = []
        for contact_id in cur.fetchall():
            cur.execute("select friendlyname from iview_Contact where id = %s", (contact_id, ))
            contacts.append(cur.fetchone()[0])
        self.contacts = '<br/>'.join(contacts)

        self.create_pdf()

        try:
            file_ = open(self.file_name, "r+")
            web.header('Content-Type', 'application/pdf', 'charset = utf-8')
            web.header('Content-disposition', 'attachment; filename = %s' % self.filepath)
            content = file_.read()
            return content
        except Exception as err:
            return err

    def create_pdf(self):
        '''
        生成pdf
        '''
        file1 = self.tick[1].encode('utf-8')
        file2 = self.tick[2].encode('utf-8')
        self.filepath = ''.join((file1, '_', file2, '.pdf'))
        self.file_name = ''.join(('/var/www/mytop/static/', self.filepath))
        file_name2 = ''.join(('/var/www/mytop/static/', file1))
        self.story = [Spacer(1, 1.5*inch)]
        self.createLineItems()
        docu = SimpleDocTemplate(file_name2)
        docu.build(self.story, onFirstPage=self.createDocument) 
        os.rename(file_name2, self.file_name)


    def time2str(self, spent):
        '''
        时间格式
        '''
        spent_m, spent_s = divmod(spent, 60)
        spent_h, spent_m = divmod(spent_m, 60)
        spent_d, spent_h = divmod(spent_h, 24)
        spent_ = ''.join((str(spent_d), '天', str(spent_h), '小时', str(spent_m), '分钟', str(spent_s), '秒'))
        return spent_


    def createDocument(self, canvas, doc):
        '''
        pdf上加内容
        '''
        ref, title, status, urgency, start, end, passed, over, spent, org, caller, agent, team, service, sub = self.tick[1:-4]

        canvas.line(1*inch, 10.6*inch, 7*inch, 10.6*inch)
        canvas.line(1*inch, 9.1*inch, 7*inch, 9.1*inch)
        canvas.line(3.9*inch, 9.1*inch, 3.9*inch, 10.6*inch)
        canvas.drawImage('/var/www/mytop/static/logo.jpg', 431, 779, 72, 33)

        ptext = "<b><font size = 14>中山云平台工单 %s</font></b>" %ref
        para = Paragraph(ptext, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 800)

        para = Paragraph(title, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 770)

        para = Paragraph("状态:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 740)
        para = Paragraph(self.statusdict[status], style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 155, 740)

        para = Paragraph("客户:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 720)
        para = Paragraph(org, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 155, 720)

        para = Paragraph("服务:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 700)
        para = Paragraph(service, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 155, 700)

        para = Paragraph("子服务:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 680)
        para = Paragraph(sub, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 155, 680)

        para = Paragraph("指派给:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 75, 660)
        para = Paragraph(agent, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 155, 660)

        para = Paragraph("起始:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 285, 740)
        para = Paragraph(str(start), style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 360, 740)

        para = Paragraph("结束:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 285, 720)
        para = Paragraph(str(end), style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 360, 720)

        para = Paragraph("持续:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 285, 700)
        if spent:
            spent_ = self.time2str(spent)
        else:
            spent_ = ''
        para = Paragraph(spent_, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 360, 700)

        para = Paragraph("过期:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 285, 680)
        if over:
            over = self.time2str(over)
        else:
            over = '0'
        para = Paragraph(over, style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 360, 680)

        para = Paragraph("优先级:", style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 285, 660)
        para = Paragraph(self.urgencys[int(urgency)], style=normalStyle)
        para.wrapOn(canvas, self.width, self.height)
        para.drawOn(canvas, 360, 660)

    def createLineItems(self):
        """
        pdf 上增加不定长的内容
        """
        desc, log, pend, solu = self.tick[-4:]

        text = desc.replace('\r\n', '<br/>')
        desc = '<br/>'.join((' ', ' ', '【描述】', '_'*85, ' ', text))
        para = Paragraph(desc, normalStyle)
        self.story.append(para)

        solution_text = solu.replace('\r\n', '<br/>')
        solu = '<br/>'.join((' ', ' ', '【解决方案】', '_'*85, ' ', solution_text))
        para = Paragraph(solu, normalStyle)
        self.story.append(para)

        header = log.replace('\r\n', '<br/>').replace('\n\n', '<br/>').replace('\n', '<br/><br/>')
        header_text = '<br/>'.join((' ', ' ', '【日志记录】', '_'*85, header))
        para = Paragraph(header_text, normalStyle)
        self.story.append(para)

        if pend:
            header = pend.replace('\r\n', '<br/>').replace('\n\n', '<br/>').replace('\n', '<br/><br/>')
            header_text = '<br/>'.join((' ', ' ', '【暂挂原因】', '_'*85, ' ', header))
            para = Paragraph(header_text, normalStyle)
            self.story.append(para)

        atts = '<br/>'.join((' ', ' ', '【附件】', '_'*85, self.atts))
        para = Paragraph(atts, normalStyle)
        self.story.append(para)

        contacts = '<br/>'.join((' ', ' ', '【联系人】', '_'*85, self.contacts))
        para = Paragraph(contacts, normalStyle)
        self.story.append(para)

urls = (
    '/atta/(.*)/(.*)', 'ATTA',
    '/pdf/(.*)', 'PDF',
)

application = web.application(urls, globals()).wsgifunc()
app = web.application(urls, globals())

if __name__ == "__main__":
    app.run()
    cur.close()
    conn.close()
