from tkinter import *
from tkinter.filedialog import askopenfilename
import pandas as pd
from tkinter import messagebox
# from pandastable import Table, TableModel

class Window(Frame):

    def __init__(self, master =None):
        Frame.__init__(self, master)

        self.master = master
        self.init_window()

    def init_window(self):

        self.master.title('1.3 Prototype')
        self.pack(fill=BOTH, expand=1)

        quitButton = Button(self, text='quit', command=self.client_exit)
        quitButton.place(x=10, y=230)

        # fileButton = Button(self, text='Browse Data Set', command=self.import_data)
        # fileButton.place(x=150, y=0)

        fileButton = Button(self, text='SBO', command=self.sbo, bg='#00a28f', fg='white')
        fileButton.place(x=240, y=150)

        fileButton = Button(self, text='CBO', command=self.cbo, bg='#00a28f', fg='white')
        fileButton.place(x=200, y=150)

        var = StringVar()
        label = Label(self.master, textvariable=var)
        label.place(x=10, y=0)

        var.set("\nInstructions:"+
                "\n1) Click at CBO or SBO button and and select the specified raw data to be calculated"+
                 "\n2) The result will be saved at file directory in excel format"+
                "\n             ")


    def client_exit(self):
        quit()

    # def import_data(self):
    #
    #     csv_file_path = askopenfilename()
    #     # print(csv_file_path)
    #     df = pd.read_excel(csv_file_path)
    #     return df

    def sbo(self):

        csv_file_path = askopenfilename()
        df = pd.read_excel(csv_file_path)

        data = df.drop(df.index[0])  # remove first row

        '''''''''''''''''''''''''''''''''''''''''''Pivot'''''''''''''''''''''''''''''''''''''''''''''''''

        data['BOVal%'] = data['BOVal%'].astype(str)  # convert to string
        data['BOQty%'] = data['BOQty%'].astype(str)
        data['CustomerPONo'] = data['CustomerPONo'].astype(str)
        data['OrdNo'] = data['OrdNo'].astype(str)
        data['VendorNo'] = data['VendorNo'].astype(str)

        pivot = data.pivot_table(index='Style', values=['BOQty','BOVal'], aggfunc='sum')  # first pivot
        pivoted = pd.DataFrame(pivot.to_records())  # flattened
        pivoted = pivoted.sort_values(by=['BOVal'], ascending=False)  # sort largest to smallest

        pivoted['Ranking'] = range(1, len(pivoted) + 1)  # Ranking

        cols = pivoted.columns.tolist()
        cols = cols[-1:] + cols[:-1]
        pivoted = pivoted[cols]
        pivoted = pivoted.set_index('Ranking')

        pivoted.loc['Total'] = pd.Series(pivoted.sum(skipna=True),
                                      index=['BOQty', 'BOVal'])
        pivoted = pivoted.sort_values(by=['BOVal'], ascending=False)
        pivoted = pivoted.round()

        '''''''''''''''''''''''''''''''''''''''''''Edited real data'''''''''''''''''''''''''''''''''''''''''''''''''
        col = data.columns.tolist()
        col = (col[22:23] + col[15:17] + col[:14] + col[17:22] + col[23:37])  # rearrange column
        data = data[col]

        data = data.sort_values(by=['BOVal'], ascending=False)  # sort value

        '''''''''''''''''''''ranking'''''''''''''''''''''''''''''''''''''''
        data['Ranking'] = range(1, len(data) + 1)  # Ranking

        columns = data.columns.tolist()
        columns = columns[-1:] + columns[:-1]
        data = data[columns]
        data = data.set_index('Ranking')

        '''''''''''''''''''''total'''''''''''''''''''''''''''''''''''''''
        data.loc['Total'] = pd.Series(data.sum(skipna=True),
                                      index=['BOQty', 'BOVal'])
        data = data.sort_values(by=['BOVal'], ascending=False)
        data = data.round()

        dates = data['SnapShotDate']
        # print(dates)
        dates = dates.iloc[2].strftime('%d%m%Y')

        sos = data['SOS']
        sos = sos[2]

        if sos == 'Kedah 2':
            sos = 'Kedah'

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter('%s SBO %s .xlsx' % (sos, dates), engine='xlsxwriter')

        # Write each dataframe to a different worksheet.
        pivoted.to_excel(writer, sheet_name='pivot')
        data.to_excel(writer, sheet_name='SBO')
        data.to_excel(writer, sheet_name=dates)

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['pivot']
        worksheet1 = writer.sheets['SBO']

        border = workbook.add_format({'border':True})
        boldv = workbook.add_format({ 'num_format': '$#,##0', 'bold': True, 'border':True})
        boldq = workbook.add_format({'num_format': '#,##0', 'bold': True, 'border':True})
        value = workbook.add_format({'num_format': '$#,##0', 'border':True})
        volume = workbook.add_format({'num_format': '#,##0', 'border':True})

        worksheet.write('D2', pivoted.iloc[0,2], boldv) #pivot worksheet
        worksheet.write('C2', pivoted.iloc[0, 1], boldq)
        worksheet.set_column('B:B', 15, border)
        worksheet.set_column('D:D', 15, value)
        worksheet.set_column('C:C', 15, volume)

        worksheet1.write('D2', data.iloc[0, 2], boldv) #sbo worksheet
        worksheet1.write('C2', data.iloc[0, 1], boldq)
        worksheet1.set_column('B:B', 15, border)
        worksheet1.set_column('D:D', 15, value)
        worksheet1.set_column('C:C', 15, volume)


        # Green fill with dark green text.
        format1 = workbook.add_format({'bg_color': '#33D2FF',
                                       'font_color': '#000000'})

        # Apply a conditional format to the cell range.

        worksheet.conditional_format('A3:D12', {'type': '2_color_scale',
                                                'min_color': '#33D2FF',
                                                'max_color': '#33D2FF'
                                                }
                                     )
        '''''''''''''''''''''''''''''''''''''''''''''''''highlight cells'''''''''''''''''''''''''''''''''''''

        worksheet1.conditional_format('A3:D12', {'type': '2_color_scale',
                                                'min_color': '#33D2FF',
                                                'max_color': '#33D2FF'})

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        messagebox.showinfo("Note", "Calculation Completed")

    def cbo(self):

        csv_file_path = askopenfilename()

        url = 'https://raw.githubusercontent.com/AimanFris/ansell/master/Stylemat.csv'
        sm = pd.read_csv(url)

        # Stylemat = askopenfilename()
        df = pd.read_excel(csv_file_path)
        # sm = pd.read_excel(Stylemat)

        df = df.drop(df.index[0])
        df.insert(loc=8, column='PH', value=['' for i in range(df.shape[0])])
        df.insert(loc=9, column='Site', value=['' for i in range(df.shape[0])])

        df['Region'] = df['Region'].fillna('"NA"')

        df['S&OP Style Aggrt'] = df['S&OP Style Aggrt'].astype(str)
        sm['Style'] = sm['Style'].astype(str)

        '''''''''''''''''''''''''''''''''''''''''''''''''Date Labelling'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        dates = df['Date_Rp']
        # print(dates)
        dates = dates.iloc[1]
        w = list(dates)
        for p in range(2):
            result = dates.find('/')
            w[result] = '-'
            dates = "".join(w)

        '''''''''''''''''''''''''''''''''''''''''''''''''VLOOKUP'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        rowcount = len(df)
        rowstyle = len(sm)

        i = 0
        j = 0
        Style = []

        for i in range(rowcount):

            for j in range(rowstyle):

                if df.iloc[i, 7] == sm.iloc[j, 0]:
                    df.iloc[i, 8] = 'Horizon'
                    df.iloc[i, 9] = sm.iloc[j, 2]

        '''''''''''''''''''''''''''''''''''''''''''''''''Pivot value'''''''''''''''''''''''''''''''''''''''''''''''''''

        table = pd.pivot_table(df[df.PH == 'Horizon'], index='S&OP Style Aggrt', columns='Region',
                               values='Net CBO Value', aggfunc='sum')
        table['Grand Total'] = table.sum(axis=1)
        table = table.sort_values(by=['Grand Total'], ascending=False)
        table = pd.DataFrame(table.to_records())
        # table['S&OP Style Aggrt'] = table['S&OP Style Aggrt'].astype(float)

        '''''''''''''''''''''ranking'''''''''''''''''''''''''''''''''''''''
        table['Ranking'] = range(1, len(table) + 1)  # Ranking

        columns = table.columns.tolist()
        columns = columns[-1:] + columns[:-1]
        table = table[columns]
        table = table.set_index('Ranking')

        '''''''''''''''''''''total'''''''''''''''''''''''''''''''''''''''
        table.loc['TOTAL'] = pd.Series(table.sum(skipna=True),
                                        index=['LAC', 'EMEA', 'APAC', '"NA"', '-', 'Grand Total'])
        table = table.round()
        table = table.sort_values(by=['Grand Total'], ascending=False)

        '''''''''''''''''''''''''''''''''''''''''''''''''Pivot volume'''''''''''''''''''''''''''''''''''''''''''''''''''

        table2 = pd.pivot_table(df[df.PH == 'Horizon'], index='S&OP Style Aggrt', columns='Region',
                                values='Net CBO Vol', aggfunc='sum')
        table2['Grand Total'] = table2.sum(axis=1)
        table2 = table2.sort_values(by=['Grand Total'], ascending=False)
        table2 = pd.DataFrame(table2.to_records())
        # table2['S&OP Style Aggrt'] = table2['S&OP Style Aggrt'].astype(float)

        '''''''''''''''''''''ranking'''''''''''''''''''''''''''''''''''''''
        table2['Ranking'] = range(1, len(table2) + 1)  # Ranking

        columns = table2.columns.tolist()
        columns = columns[-1:] + columns[:-1]
        table2 = table2[columns]
        table2 = table2.set_index('Ranking')

        '''''''''''''''''''''total'''''''''''''''''''''''''''''''''''''''
        table2.loc['TOTAL'] = pd.Series(table2.sum(skipna=True),
                                       index=['LAC', 'EMEA', 'APAC', '"NA"', '-', 'Grand Total'])
        table2 = table2.round()
        table2 = table2.sort_values(by=['Grand Total'], ascending=False)

        '''''''''''''''''''''''''''''''''''''''''''''''''Export to excel'''''''''''''''''''''''''''''''''''''''''''''
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter('CBO %s .xlsx' % dates, engine='xlsxwriter')

        # Write each dataframe to a different worksheet.
        table.to_excel(writer, sheet_name='CBO_value')
        table2.to_excel(writer, sheet_name='CBO_volume')
        df.to_excel(writer, sheet_name=dates)
        sm.to_excel(writer, sheet_name='StyleMat')

        # Get the xlsxwriter workbook and worksheet objects.
        workbook = writer.book
        worksheet = writer.sheets['CBO_value']
        worksheet2 = writer.sheets['CBO_volume']

        bold = workbook.add_format({'num_format': '0', 'bold':True, 'border':True})

        boldq = workbook.add_format({'num_format': '#,##0', 'bold': True, 'border':True})
        boldvr = workbook.add_format({'num_format': '$#,##0', 'bold': True})
        boldqr = workbook.add_format({'num_format': '#,##0', 'bold': True})
        value = workbook.add_format({'num_format': '$#,##0', 'border':True})
        volume = workbook.add_format({'num_format': '#,##0', 'border':True})

        worksheet.set_column('C:H', 10, value)
        worksheet.set_column('A:A', 8, boldq)
        worksheet.set_column('B:B', 15, bold)
        worksheet.set_row(1, 14.5, boldvr)

        worksheet2.set_column('C:H', 10, volume)
        worksheet2.set_column('A:A', 8, boldq)
        worksheet2.set_column('B:B', 15, bold)
        worksheet2.set_row(1, 14.5, boldqr)

        # Green fill with dark green text.
        format1 = workbook.add_format({'bg_color': '#33D2FF',
                                       'font_color': '#000000',
                                       'bold': 1,
                                       })

        # Apply a conditional format to the cell range.

        worksheet.conditional_format('A3:H12', {'type': '2_color_scale',
                                                'min_color': '#33D2FF',
                                                'max_color': '#33D2FF',
                                                }
                                     )

        worksheet.conditional_format('A3:H12', {'type': 'blanks',
                                                'format': format1})

        # Apply a conditional format to another cell range.

        worksheet2.conditional_format('A3:H12', {'type': '2_color_scale',
                                                'min_color': '#33D2FF',
                                                'max_color': '#33D2FF',
                                                }
                                     )

        worksheet2.conditional_format('A3:H12', {'type': 'blanks',
                                                'format': format1})

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        messagebox.showinfo("Note", "Calculation Completed")


root = Tk()
root.geometry('470x300')

app = Window(root)

root.mainloop()
