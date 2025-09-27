import pandas as pd
from docxtpl import DocxTemplate,InlineImage
from docxcompose.composer import Composer
import os
import matplotlib.pyplot as plt
import io
import base64
from docx.shared import Mm
import matplotlib.pyplot as plt
import textwrap
from datetime import date





def script(school_name,principal_name,goals_path,
            player_data_path,
            summary_path,
            educator_path,
            family_path):
    document = DocxTemplate("part1Template[1].docx")
    documentkg1=DocxTemplate('part2Template[1].docx')
    documentkg2=DocxTemplate('part2Template[1].docx')
    document1=DocxTemplate('part2Template[1].docx')
    document2=DocxTemplate('part2Template[1].docx')
    document3=DocxTemplate('part2Template[1].docx')
    document4=DocxTemplate('part2Template[1].docx')
    document5=DocxTemplate('part2Template[1].docx')
    document6=DocxTemplate('part2Template[1].docx')
    document7=DocxTemplate('part2Template[1].docx')
    document8=DocxTemplate('part2Template[1].docx') 
    document9 = DocxTemplate('part3.docx')
    today_date = date.today()
    document_list=[]
    json_object = {}
    object_list=[]
    
    Goals_data=pd.read_excel(goals_path)
    player_data = pd.read_excel(player_data_path)
    School_Summary = pd.read_excel(summary_path)
    Educator_data=pd.read_excel(educator_path)
    Family_data=pd.read_excel(family_path)

    json_object['p_School']=school_name
    json_object['p10']=principal_name
    total_sum=School_Summary['Est Math Problems Solved'].sum()
    json_object['p1']=str(total_sum)

    average_sums_solved = (School_Summary['Est Math Problems Solved'].sum())//School_Summary['Students'].sum()
    json_object['p2']=str(average_sums_solved)

    Teams_with_1000 = (School_Summary['Stickers / Student']>1000).sum()
    total_teams=(School_Summary['Stickers / Student']).count()
    json_object['p4']=total_teams
    json_object['p3']=str(Teams_with_1000)
    # Representation of top 5
    top_5_teams = School_Summary.sort_values(by='Est Math Problems Solved', ascending=False).head(5)[['Name', 'Class Name']]


    json_object['p_18']=InlineImage(document,createTable(top_5_teams,20,5,0.75,1,35,False),width=Mm(100))

    top_5_player = player_data.sort_values(by='Stickers', ascending=False).head(5)[['User ID', 'First Name', 'Last Name','Display Name']]
    json_object['p_19']=InlineImage(document,createTable(top_5_player,20,5,0.75,1,35,False),width=Mm(100))

    worksheets_saved=total_sum//20
    json_object['p7']=str(worksheets_saved)

    time_saved=worksheets_saved//80
    json_object['p8']=str(time_saved)
    json_object['p11']=today_date
    json_object['p_12']="Chota lund"

   

    assissment_report_kg1=School_Summary[School_Summary['Class Name'].str.startswith('KG-1')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_kg1 = assissment_report_kg1.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_kg1.empty: 
        temp={}
        temp['p_24']=InlineImage(documentkg1,createTable(sorted_assissment_report_kg1,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='KG-1'
        temp['p_25']='Early'
        object_list.append(temp)
        
    Goal_Index_kg1=Goals_data[Goals_data['Class Name'].str.startswith('KG-1')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_kg1 = Goal_Index_kg1.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_kg1.empty:
        object_list[-1]['p_26']=InlineImage(documentkg1,createTable(sorted_Goal_Index_kg1,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        documentkg1.render(object_list[-1])
        document_list.append(documentkg1)
        

    assissment_report_kg2=School_Summary[School_Summary['Class Name'].str.startswith('KG-2')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_kg2 = assissment_report_kg2.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_kg2.empty: 
        temp={}
        temp['p_24']=InlineImage(documentkg2,createTable(sorted_assissment_report_kg2,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='KG-2'
        temp['p_25']='Early'
        object_list.append(temp)

    Goal_Index_kg2=Goals_data[Goals_data['Class Name'].str.startswith('KG-2')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_kg2 = Goal_Index_kg2.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_kg2.empty:
        object_list[-1]['p_26']=InlineImage(documentkg2,createTable(sorted_Goal_Index_kg2,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        documentkg2.render(object_list[-1])
        document_list.append(documentkg2)
    
    assissment_report_1=School_Summary[School_Summary['Class Name'].str.startswith('1')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_1 = assissment_report_1.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_1.empty: 
        temp={}
        temp['p_24']=InlineImage(document1,createTable(sorted_assissment_report_1,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='1'
        temp['p_25']='Early'
        object_list.append(temp)

    Goal_Index_1=Goals_data[Goals_data['Class Name'].str.startswith('1')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_1 = Goal_Index_1.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_1.empty:
        object_list[-1]['p_26']=InlineImage(document1,createTable(sorted_Goal_Index_1,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document1.render(object_list[-1])
        document_list.append(document1)

    assissment_report_2=School_Summary[School_Summary['Class Name'].str.startswith('2')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_2 = assissment_report_2.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_2.empty: 
        temp={}
        temp['p_24']=InlineImage(document2,createTable(sorted_assissment_report_2,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='2'
        temp['p_25']='Early'
        object_list.append(temp)

    Goal_Index_2=Goals_data[Goals_data['Class Name'].str.startswith('2')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_2 = Goal_Index_2.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_2.empty:
        object_list[-1]['p_26']=InlineImage(document2,createTable(sorted_Goal_Index_2,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document2.render(object_list[-1])
        document_list.append(document2)

    assissment_report_3=School_Summary[School_Summary['Class Name'].str.startswith('3')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_3 = assissment_report_3.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_3.empty: 
        temp={}
        temp['p_24']=InlineImage(document3,createTable(sorted_assissment_report_3,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='3'
        temp['p_25']='Elementary'
        object_list.append(temp)

    Goal_Index_3=Goals_data[Goals_data['Class Name'].str.startswith('3')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_3 = Goal_Index_3.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_3.empty:
        object_list[-1]['p_26']=InlineImage(document3,createTable(sorted_Goal_Index_3,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document3.render(object_list[-1])
        document_list.append(document3)
    
    assissment_report_4=School_Summary[School_Summary['Class Name'].str.startswith('4')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_4 = assissment_report_4.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_4.empty: 
        temp={}
        temp['p_24']=InlineImage(document4,createTable(sorted_assissment_report_4,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='4'
        temp['p_25']='Elementary'
        object_list.append(temp)

    Goal_Index_4=Goals_data[Goals_data['Class Name'].str.startswith('4')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_4 = Goal_Index_4.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_4.empty:
        object_list[-1]['p_26']=InlineImage(document4,createTable(sorted_Goal_Index_4,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document4.render(object_list[-1])
        document_list.append(document4)

    assissment_report_5=School_Summary[School_Summary['Class Name'].str.startswith('5')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_5 = assissment_report_5.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_5.empty: 
        temp={}
        temp['p_24']=InlineImage(document5,createTable(sorted_assissment_report_5,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='5'
        temp['p_25']='Intermediate'
        object_list.append(temp)

    Goal_Index_5=Goals_data[Goals_data['Class Name'].str.startswith('5')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_5 = Goal_Index_5.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_5.empty:
        object_list[-1]['p_26']=InlineImage(document5,createTable(sorted_Goal_Index_5,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document5.render(object_list[-1])
        document_list.append(document5)

    assissment_report_6=School_Summary[School_Summary['Class Name'].str.startswith('6')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_6 = assissment_report_6.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_6.empty: 
        temp={}
        temp['p_24']=InlineImage(document6,createTable(sorted_assissment_report_6,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='6'
        temp['p_25']='Intermediate'
        object_list.append(temp)

    Goal_Index_6=Goals_data[Goals_data['Class Name'].str.startswith('6')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_6 = Goal_Index_6.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_6.empty:
        object_list[-1]['p_26']=InlineImage(document6,createTable(sorted_Goal_Index_6,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document6.render(object_list[-1])
        document_list.append(document6)


    assissment_report_7=School_Summary[School_Summary['Class Name'].str.startswith('7')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_7 = assissment_report_7.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_7.empty: 
        temp={}
        temp['p_24']=InlineImage(document7,createTable(sorted_assissment_report_7,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='7'
        temp['p_25']='Advanced'
        object_list.append(temp)

    Goal_Index_7=Goals_data[Goals_data['Class Name'].str.startswith('7')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_7 = Goal_Index_7.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_7.empty:
        object_list[-1]['p_26']=InlineImage(document7,createTable(sorted_Goal_Index_7,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document7.render(object_list[-1])
        document_list.append(document7)

    assissment_report_8=School_Summary[School_Summary['Class Name'].str.startswith('8')][['Class Name','Teacher','Std','Students','Stickers','Stickers / Student','Teacher User Id','Est Time On Task (Hours)','Est Math Problems Solved']]
    sorted_assissment_report_8 = assissment_report_8.sort_values(by='Class Name', ascending=True)
    if not sorted_assissment_report_8.empty: 
        temp={}
        temp['p_24']=InlineImage(document8,createTable(sorted_assissment_report_8,40,10,1,1,28,True),width=Mm(180),height=Mm(65))
        temp['p_20']='8'
        temp['p_25']='Advanced'
        object_list.append(temp)

    Goal_Index_8=Goals_data[Goals_data['Class Name'].str.startswith('8')][['Class Name','Goals Index (out of 100)','Activity (out of 25)','Fact Fluency (out of 25)']]
    sorted_Goal_Index_8 = Goal_Index_8.sort_values(by='Class Name', ascending=True)
    if not sorted_Goal_Index_8.empty:
        object_list[-1]['p_26']=InlineImage(document8,createTable(sorted_Goal_Index_8,40,10,1,1,35,True),width=Mm(160),height=Mm(50))
        document8.render(object_list[-1])
        document_list.append(document8)

    educator_players=Educator_data['Player'].count()
    json_object['X']=str(educator_players)
    educator_problems=(Educator_data['Sticker Count'].sum())*3
    json_object['p5']=str(educator_problems)

    family_players=Family_data['Player'].count()
    json_object['Y']=str(family_players)

    family_problems=(Family_data['Sticker Count'].sum())*3
    json_object['p6']=str(family_problems)

    document.render(json_object)
    document9.render({})
    document_list.append(document9)
    if document_list:
        master=document
        composer=Composer(master)
        for doc_to_append in document_list[0:]:
            composer.append(doc_to_append)
        
        # 1. Create a memory buffer
        final_doc_buffer = io.BytesIO()

        # 2. Save the composed document to the buffer instead of a file
        composer.save(final_doc_buffer)

        # 3. Important: Go back to the start of the buffer
        final_doc_buffer.seek(0)

        # 4. Return the buffer so Flask can use it
        return final_doc_buffer

    print(document_list)


def createTable(DataSet,h,w,h1,w1,f,wrap):
    fig, ax = plt.subplots(figsize=(h,w))  # Adjust figure size as needed
    if wrap:
        wrapped_labels = ['\n'.join(textwrap.wrap(label, width=12)) for label in DataSet.columns]
    else:
        wrapped_labels=DataSet.columns
    
    # Hide axes
    ax.axis('off')

    # Create the table
    # You can customize colLabels, bbox, and cellLoc for appearance
    table = ax.table(cellText=DataSet.values,
                    colLabels=wrapped_labels,
                    cellLoc='center',
                    loc='center',bbox=[0, 0, h1, w1])

    # Style the table
    table.auto_set_font_size(False)
    table.set_fontsize(f)
    table.scale(1, 1) # Adjust column width and row height
    

    # You can add more styling here, e.g., for headers or specific cells
    for (row, col), cell in table.get_celld().items():
        cell.set_edgecolor('black')
        if row == 0:  # Header row
            cell.set_facecolor("#A5B1FF")
            cell.set_text_props(color='black', weight='bold',wrap=True)
        else:
            cell.set_facecolor("#FFFFFF") # Alternating row colors could be added here

    plt.tight_layout()

    # Save the table as an image
    buf = io.BytesIO()
    plt.savefig(buf, format="png")
    buf.seek(0)

    # 3. Encode as base64 string (safe for JSON)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")

    img_bytes = base64.b64decode(img_base64)
    img_stream=io.BytesIO(img_bytes)
    # 4. Build JSON object
    return img_stream

# Display the plot if running in an int

