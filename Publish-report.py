import pandas as pd
import os

path = '../PUBLISH REPORT'
PICs = 'Abid'
submit_date = '3/01/2024'
success_date = '4/01/2024'

class publish_report:
    def __init__(self, path_folder, PIC, submitDate, successDate):
        self.path_ticket = path_folder
        self.pics = PIC
        self.submitDate = submitDate
        self.successDate = successDate
        self.tickets_list = []
        self.Channels = {
            'DF-PSA' :  'DF PSA',
            'DF-SS' :  'DF SS',
            'LAZADA-SS' :  'LAZADA SS',
            'MAPCLUB-PSA' :  'MAPCLUB PSA',
            'MAPCLUB-SS' :  'MAPCLUB SS',
            'SHOPEE-SS' :  'SHOPEE SS',
            'DF-FOOTLOCKER' :  'DF FOOTLOCKER',
            'DF-SKECHER' :  'DF SKECHER',
            'DF-REEBOK' :  'DF REEBOK',
            'SHOPEE-SKECHER' : 'SHOPEE SKECHER',
            'SHOPEE-ASTEC' :  'SHOPEE ASTEC',
            'SHOPEE-REEBOK' : 'SHOPEE REEBOK',
            'SHOPEE-CROCS' : 'SHOPEE CROCS'
        }
    
    def _data_channel(self,file,channel,ticket):
        df_channel = file
        headers = list(df_channel.columns)
        SPU_code = df_channel['SPU'].tolist()
        SKU_code = df_channel['SKU Code'].tolist()
        Brand_name =  df_channel['Brand Name'].tolist()
        
        Channels = channel
        if Channels in ["DF SS", "MAPCLUB PSA", "MAPCLUB SS"]:
            # Channel_category = df_channel.drop(df_channel.index[-1])
            Channel_category = df_channel[f'{headers[-1]}'].tolist()
        else:
            Channel_category = '-'
            
        No_ticket = ticket
        Pics = self.pics
        submit_dates = self.submitDate
        success_dates = self.successDate
        
        # if ['remark','remarks','Remark','Remarks'] in headers:
        if 'Remarks' in headers or 'remarks' in headers or headers[1]=='Unnamed: 1':
            remarks = df_channel[f'{headers[0]}'].tolist()
            details = df_channel[f'{headers[1]}'].tolist()
        else:
            remarks = 'publish success'
            details = ''
              
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
        for ticket in self.tickets_list:
            path_channel = os.listdir(f'./{ticket}')   
            with pd.ExcelWriter(f'AOS-{ticket} publish report.xlsx') as writer:
                for files in path_channel:
                    for chan in self.Channels:
                        if chan in files:
                            channel = ''
                            file = pd.read_excel(f'./{ticket}/{files}',sheet_name=0)
                            channel = self.Channels[chan]
                            data_frame = create_publish_report._data_channel(file,channel,f'AOS-{ticket}')
                            data_frame.to_excel(writer,sheet_name=channel,index=False,header=True)
                        else:
                            continue
                print(f'AOS-{ticket} publish report Created')
        return None                
       
    
create_publish_report = publish_report(path, PICs, submit_date, success_date)
create_publish_report.Create()
