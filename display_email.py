#!/usr/bin/env python
# -*- coding: utf-8 -*-

import win32com.client
import os

sender = ''''''
to = ''''''
cc = ''''''
name = ''


def Mail_Create(sender, to, cc, name, attach_path='H:\\draw_out'):
	outlook = win32com.client.Dispatch('Outlook.Application')
	attach_files = []
	for dirt in os.listdir(attach_path):
		attach_files.insert(1000, dirt)
	attach_files.sort()
	first = attach_files[0]
	last = attach_files[-1][-4:]
	one = outlook.CreateItem(0)
	one.SentOnBehalfOfName  = sender
	one.CC = cc
	one.To = to
	one.ReadReceiptRequested = True
	one.Subject = 'Sealink - Antrag auf Garantiezi ehung ({0})/Guarantee Draw Requests ({1})'.format(first, first)
	one.Body = u"Sehr geehrte Damen und Herren,\n\nAnbei die Genehmigungsanfragen für die Ziehung der Garantie Nr. {0} (über die nächsten Emails verteilt, werden Sie die Garantieziehungen bis Nr. {1} erhalten) und die dazugehörigen Trustee-Reports, die als unterstützende Dokumentation für die Anträge auf Garantieziehung dienen.\n\nAnbei finden Sie ein Excel File, was zeigt die Differenz zwischen der Ursprünglichen Verlusthöhe und der Endgültigen Verlusthöhe.\nBitte bestätigen Sie den Erhalt dieser Email.\n\nMit freundlichen Grüßen,\n{2}\n\n--------\n\nDear all,\n\nPlease find attached the guarantee draw requests for guarantee draws # {3} (over the next few emails you will receive the requests up to #{4}) and the respective trustee reports which serve as supporting documentation for each of the guarantee draw requests.\nI have included an excel file showing the difference between the notified amount and the final draw amount.\n\nPlease confirm the receipt of this email.\n\nBest regards,\n\n{5}".format(first, last, name, first, last, name)
	for a_f in os.listdir(os.path.join(attach_path, first)):
		att = os.path.join(attach_path, first, a_f)
		one.Attachments.Add(att)
	attach_files.remove(first)
	one.Close(0)
	for email_num in attach_files:
		msg = outlook.CreateItem(0)
		msg.SentOnBehalfOfName  = sender
		msg.CC = cc
		msg.To = to
		msg.ReadReceiptRequested = True
		msg.Subject = 'Sealink - Antrag auf Garantiezi ehung ({0})/Guarantee Draw Requests ({1})'.format(email_num, email_num)
		msg.Body = u"Wie in meiner vorigen Email angekündigt, anbei die Garantieziehungen Nr. {0}.\n\nBitte bestätigen Sie den Erhalt dieser Email.\n\nMit freundlichen Grüßen,\n\n{1}\n\n--------\n\nAs announced in my previous Email, please find attached guarantee draws #{2}\n\nPlease confirm the receipt of this Email.\n\nThank you,\n\n{3}".format(email_num, name, email_num, name)
		for a_p in os.listdir(os.path.join(attach_path, email_num)):
			attachment = os.path.join(attach_path, email_num, a_p)
			msg.Attachments.Add(attachment)
		msg.Close(0)


Mail_Create(sender,to,cc,name)