from datetime import date

from PyQt6.QtWidgets import QMessageBox
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas
import pandas as pd

from error_msg import error_msg


def page_setup(canvas, doc):
    title = "Report"
    pageinfo = "This document is strictly private and confidential"
    h, w = A4
    canvas.saveState()
    canvas.setFont('Times-Roman', 24)
    canvas.drawCentredString(w / 2.0, h - 50 * mm, title)
    canvas.setFont('Times-Roman', 9)

    canvas.drawCentredString(w / 2, 16 * mm, "%s" % pageinfo)
    canvas.drawCentredString(20 * mm, 16 * mm, "%s" % date.today())

    # removed functionality for line draw as it clashed with the table
    #canvas.setStrokeColorRGB(0, 119 / 255, 181 / 255)
    #canvas.setLineWidth(2)
    #outer_box = [(8 * mm, 8 * mm, w - 8 * mm, 8 * mm), (8 * mm, h - 8 * mm, w - 8 * mm, h - 8 * mm),
    #(8 * mm, h - 8 * mm, 8 * mm, 8 * mm), (w - 8 * mm, h - 8 * mm, w - 8 * mm, 8 * mm)]
    #canvas.lines(outer_box)  # outer box draw

    canvas.drawImage("required/better_logo.png", w - 78 * mm, h - 30 * mm, width=62.5 * mm, height=15 * mm, mask='auto')

    canvas.restoreState()


def dataframe_formating(csv_to_table, step):
    filtered = csv_to_table.loc[
        (csv_to_table['step'] == step), ["client", "campaign", "sent", "delivered", "opened", "opened_rate",
                                         "responded",
                                         "responded_rate", "interested_yes", "bounced",
                                         "bounce_rate", "opt_out", "opt_out_rate"]]
    client = filtered['client'].values[0]
    sum_cols = ["sent", "delivered", "opened", "responded", "interested_yes", "bounced", "opt_out"]
    avg_cols = ["opened_rate", "responded_rate", "bounce_rate", "opt_out_rate"]
    summed = filtered[sum_cols].sum(axis=0)
    avgd = filtered[avg_cols].mean(axis=0)

    concat = pd.concat([summed, avgd], axis="columns")
    filled = concat[0].fillna(concat[1])

    transposed = filled.reset_index()
    transposed.columns = ['Header', 'Values']
    final = transposed.set_index('Header').T

    final.rename(columns={'sent': 'Sent', 'delivered': 'Delivered', 'opened': 'Opened', 'responded': 'Responded',
                          'interested_yes': 'Interested', 'bounced': 'Bounced', 'opt_out': 'Unsubscribed',
                          'opened_rate': 'Open Rate', 'responded_rate': 'Response Rate',
                          'bounce_rate': 'Bounce Rate', 'opt_out_rate': 'Unsubscribed Rate'}, inplace=True)

    format_mapping = {'Sent': '{:,.0f}', 'Delivered': '{:,.0f}', 'Opened': '{:,.0f}', 'Responded': '{:,.0f}',
                      'Interested': '{:,.0f}', 'Bounced': '{:,.0f}', 'Unsubscribed': '{:,.0f}',
                      'Open Rate': '{:,.2f}%',
                      'Response Rate': '{:,.2f}%', 'Bounce Rate': '{:,.2f}%', 'Unsubscribed Rate': '{:,.2f}%'}

    for key, value in format_mapping.items():
        final[key] = final[key].apply(value.format)
    if step == 1:
        final.insert(0, 'Client', client)
        final.insert(0, 'Step', 'Initial Email')

    elif step == 2:
        final.insert(0, 'Client', client)
        final.insert(0, 'Step', '1st Follow Up')
        final.to_numpy()

    elif step == 3:
        final.insert(0, 'Client', client)
        final.insert(0, 'Step', '2nd Follow Up')
        final.to_numpy()

    else:
        error_msg("somethings gone horribly wrong")

    return final


def table_creation(csv_to_table):  # creates stylised tables
    h, w = A4
    datalist = [csv_to_table.columns.tolist()] + csv_to_table.values.tolist()
    table = Table(datalist, hAlign=TA_CENTER, style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                                                     ('VALIGN', (0, 0), (7, 0), 'MIDDLE'),
                                                     ('ALIGN', (0, 0), (7, 0), 'CENTER'),
                                                     ('ALIGN', (0, 0), (7, 0), 'CENTER'),
                                                     ('BACKGROUND', (0, 0), (-1, 0), colors.rgb2cmyk(0, 119 / 255, 181 / 255)),
                                                     ('BACKGROUND', (0, 0), (0, -1), colors.rgb2cmyk(0, 119 / 255, 181 / 255)),
                                                     ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                                                     ('TEXTCOLOR', (0, 0), (0, -1), colors.white)], cornerRadii=[8, 8, 8, 8])

    return table


def generation_message(save_location):  # message box to tell user the pdf has been generated
    msg_box = QMessageBox()
    msg_box.setWindowTitle("Generation Status")
    msg_box.setText('PDF Generated, saved to ' + save_location + '.pdf')
    msg_box.exec()


def pdf(filtered_csv, save_location):
    h, w = A4
    elements = [Spacer(w, 40 * mm)]
    final_dataframe = pd.DataFrame()
    for i in range(1, 4):
        formatted_df = dataframe_formating(filtered_csv, i)
        if i == 1:
            final_dataframe = formatted_df
        else:
            no_header_df = pd.DataFrame(formatted_df.values, columns=final_dataframe.columns)
            final_dataframe = pd.concat([final_dataframe, no_header_df], ignore_index=True)
    elements.append(table_creation(final_dataframe))

    doc = SimpleDocTemplate(save_location + '.pdf', pagesize=landscape(A4))
    doc.build(elements, onFirstPage=page_setup)
    generation_message(save_location)
