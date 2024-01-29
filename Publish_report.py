import pandas as pd
import os
from colorama import Fore, Style
from openpyxl.styles import Font, Side, Border 

path = '../_PUBLISH REPORT'
folder_dest = 'publish_report'
PICs = 'Fadhli'
submit_date ='23-Jan-2024' #d/m/yyyy
success_date = '24-Jan-2024'
Channels_ticket = {
            'LAZADA-SS' :  'LAZADA SS',
            'MAPCLUB-PSA' :  'MAPCLUB PSA',
            'MAPCLUB-SS' :  'MAPCLUB SS',
            'DF-SS' :  'DF SS',
            'DF-PSA' :  'DF PSA',
            'DF-FOOTLOCKER' :  'DF FOOTLOCKER',
            'DF-SKECHER' :  'DF SKECHER',
            'DF-REEBOK' :  'DF REEBOK',
            'DF-CROCS' :  'DF CROCS',
            'SHOPEE-SS' :  'SHOPEE SS',
            'SHOPEE-SKECHERS' : 'SHOPEE SKECHERS',
            'SHOPEE-ASTEC' :  'SHOPEE ASTEC',
            'SHOPEE-CROCS' : 'SHOPEE CROCS',
            'SHOPEE-REEBOK' : 'SHOPEE REEBOK',
            'SHOPEE-CON' : 'SHOPEE CON',
            'MONO-CON' : 'MONO CON'
        }
Channels = {key.lower(): value for key,value in Channels_ticket.items()}

class publish_report:
    def __init__(self, path_folder, PIC, submitDate, successDate):
        self.path_ticket = path_folder
        self.pics = PIC
        self.submitDate = submitDate
        self.successDate = successDate
        self.tickets_list = []
        # self.headerzz = []
        
    
    def _data_channel(self,file,channel,ticket):
        df_channel = file
        headers = list(df_channel.columns)
        SPU_code = df_channel['SPU'].tolist()
        SKU_code = df_channel['SKU Code'].tolist()
        Brand_name =  df_channel['Brand Name'].tolist()
        have_categoryChan = ["DF SS", "MAPCLUB PSA", "MAPCLUB SS",
                            "LAZADA SS", "SHOPEE REEBOK", "SHOPEE CROCS",
                            "SHOPEE ASTEC", "SHOPEE SKECHERS", "SHOPEE SS",
                            "SHOPEE CON" ]
        Channels = channel
        if Channels in have_categoryChan:
            header_nmcat = 'category_'
            if header_nmcat in headers[-1]:
                Channel_category = df_channel[f'{headers[-1]}'].tolist()
            else:
                for i in headers:
                    if header_nmcat in i:
                        Channel_category = df_channel[i].tolist()
        else:
            Channel_category = '-'
              
        No_ticket = ticket
        Pics = self.pics
        submit_dates = self.submitDate
        # success_dates = self.successDate
        
        # if ['remark','remarks','Remark','Remarks'] in headers:
        if 'Remarks' in headers or 'remarks' in headers or headers[1]=='Unnamed: 1':
            remarks = df_channel[f'{headers[0]}'].tolist()
            details = df_channel[f'{headers[1]}'].tolist()
            try:
                success_dates = [self.successDate if i.lower() == 'publish success' else '' for i in remarks]
            except:
                success_dates = [self.successDate if i == 'Publish success' else '' for i in remarks]
        else:
            remarks = 'Publish success'
            details = ''
            success_dates = self.successDate
    
        data_channel = pd.DataFrame({
            'SPU': SPU_code,
            'SKU': SKU_code,
            'Brand Name' : Brand_name,
            'Channel' : Channels,
            'Channel category' : Channel_category,
            'No. Ticket': No_ticket,
            'PIC': Pics,
            'Submit Date': submit_dates,
            'Success Date': success_dates,
            'Remarks': remarks,
            'Details': details
        })
        # self.headerzz = data_channel.columns.tolist()
        print(headers[-1])   
        return data_channel
      
    def len_data(data,count):
        data = data
        data_list = []
        for i in range(len(count)):
            data_list.append(data)
        return data_list
    
    def _tickets(self):
        tickets =  os.listdir(self.path_ticket)
        for tick in tickets:
            full_path = os.path.join(self.path_ticket,tick)
            if os.path.isdir(full_path):
                self.tickets_list.append(tick)
        return self.tickets_list
    
    def Create(self):
        create_publish_report._tickets()
        if folder_dest not in self.tickets_list:
            os.mkdir(folder_dest)
            
        print(Fore.GREEN + 'Creating publish report...' + Style.RESET_ALL )
        for ticket in self.tickets_list:
            path_channel = os.listdir(f'./{ticket}')
            path_channel = [i.lower() for i in path_channel]
            if ticket == '__pycache__' or ticket == folder_dest:
                continue
            try:
                with pd.ExcelWriter(f'{folder_dest}/AOS-{ticket} publish report.xlsx', engine='openpyxl') as writer:
                    for files in path_channel:
                        for chan in Channels:
                            if chan in files:
                                channel = ''
                                file = pd.read_excel(f'./{ticket}/{files}',sheet_name=0)
                                channel = Channels[chan]
                                data_frame = create_publish_report._data_channel(file,channel,f'AOS-{ticket}')
                                data_frame = data_frame.drop(data_frame[data_frame['SPU'].str.contains('SPU')].index)
                                #delete row that contains 'SPU', to fix the template that have 2 row headers 
                                data_frame.to_excel(writer,sheet_name=channel,index=False,header=True)
                                workbook = writer.book
                                worksheet = writer.sheets[channel]
                                
                                #font bold and non-bordered header
                                header_font = Font(bold=True)
                                border = Border()
                                
                                for cell in worksheet[1]:
                                    cell.font = header_font
                                    cell.border = border
                            else:
                                continue
                print(f'AOS-{ticket} publish report Created')
            except:
                print(Fore.RED + f'AOS-{ticket} publish report failed to Create'+ Style.RESET_ALL )
                continue
        print(Fore.GREEN + 'Process completed!' + Style.RESET_ALL )
        return None                
       
create_publish_report = publish_report(path, PICs, submit_date, success_date)
# create_publish_report.Create()
