import pandas as pd
import tkinter as tk
import xlsxwriter

from tkinter import filedialog


def convert():
    filename = filedialog.askopenfilename(filetypes=([('Xlsx files', '*.xlsx'),('All files', '*.*')]))
    batch_number = batchnum_entry.get()
    pickup_date = pickup_date_entry.get()

    my_sheet = 'For Pickup'
    df = pd.read_excel(filename, sheet_name = my_sheet)
    df.columns = [
        'index',
        'date',
        'usana_id',
        'usana_name',
        'type',
        'sales_order',
        'product_name',
        'quantity',
        'total_price',
        'total',
        'delivery_address',
        'sf',
        'each',
        'add',
        'total',
        'tracking_number'
    ]

    df.dropna(subset=['sales_order'], inplace=True)
    length = int(df.shape[0])
    
    so_list = []
    assoc_name_list = []
    usana_id_list = []
    type_list = []
    product_list = []

    # Fetch data from original masterlsit
    for rec in range(length):
        sales_order = df["sales_order"].values[rec]
        usana_name = df["usana_name"].values[rec]
        usana_id = df["usana_id"].values[rec]
        usana_type = df["type"].values[rec]
        product_name = df["product_name"].values[rec]

        if rec == 0:
            continue
        else:
            so_list.append(sales_order)
            assoc_name_list.append(usana_name)
            usana_id_list.append(usana_id)
            type_list.append(usana_type)
            product_list.append(product_name)

    # Convert into another masterlist
    file_directory = filedialog.askdirectory()
    converted_masterlist = file_directory + '/Batch-{}_masterlist.xlsx'.format(batch_number)
    workbook = xlsxwriter.Workbook(converted_masterlist)

    pickup_masterlist = workbook.add_worksheet('PickupList')
    product_masterlist = workbook.add_worksheet('ProductList')

    # Formats
    title_format = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri (Body)',
        'font_size': 26,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    header_format = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri (Body)',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FFD966',
        'border': 1
    })
    so_format = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri (Body)',
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    content_format = workbook.add_format({
        'bold': False,
        'font_name': 'Calibri (Body)',
        'font_size': 12,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    name_format = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri (Body)',
        'font_size': 14,
        'align': 'left',
        'valign': 'vcenter',
        'left': 2,
        'bottom': 1
    })
    product_format = workbook.add_format({
        'bold': False,
        'font_name': 'Calibri (Body)',
        'font_size': 14,
        'align': 'left',
        'valign': 'vcenter',
        'right': 2,
        'bottom': 1,
        'left': 1
    })
    bottom_border_format = workbook.add_format({
        'bottom': 2
    })
    top_border_format = workbook.add_format({
        'top': 2
    })

    # Pickup Masterlist Sheet
    pickup_masterlist.merge_range('A1:D1','{} - DAVAO WILL CALL PICK UP'.format(pickup_date),title_format)
    pickup_masterlist.write('A2','Order Number',header_format)
    pickup_masterlist.write('B2','Associate\'s Name',header_format)
    pickup_masterlist.write('C2','USANA ID',header_format)
    pickup_masterlist.write('D2','Type',header_format)

    index = 3
    for rec in range(len(so_list)):
        pickup_masterlist.write('A' + str(index), so_list[rec], so_format)
        pickup_masterlist.write('B' + str(index), assoc_name_list[rec], content_format)
        pickup_masterlist.write('C' + str(index), usana_id_list[rec], content_format)
        pickup_masterlist.write('D' + str(index), type_list[rec], content_format)
        index += 1

    pickup_masterlist.set_column('A:A', 17.00)
    pickup_masterlist.set_column('B:B', 26.67)
    pickup_masterlist.set_column('C:C', 17.00)
    pickup_masterlist.set_column('D:D', 17.00)

    # Product Masterlist Sheet
    index2 = 2

    for rec in range(len(so_list)):
        products_rec = product_list[rec].split(",")
        products_split = [products_rec[i * 5:(i + 1) * 5] for i in range((len(products_rec) + 5 - 1) // 5 )]
        loop_count = 1

        index_start = index2
        for by_five in products_split:
            product_string = ''

            for product in by_five:
                product_string += '{},'.format(product)

            product_masterlist.write('B' + str(index2), product_string, product_format)
            index2 += 1
            loop_count += 1

        index_end = index2

        if index_start != index_end-1:
            range_merge = 'A{}:A{}'.format(index_start, index_end-1)
            product_masterlist.merge_range(range_merge, assoc_name_list[rec],name_format)
        else:
            product_masterlist.write('A' + str(index2-1), assoc_name_list[rec], name_format)

    product_masterlist.set_column('A:A', 30.00)
    product_masterlist.set_column('B:B', 81.00)

    product_masterlist.write('A1', None, bottom_border_format)
    product_masterlist.write('B1', None, bottom_border_format)
    product_masterlist.write('A' + str(index2), None, top_border_format)
    product_masterlist.write('B' + str(index2), None, top_border_format)

    workbook.close()

    root.destroy()


if __name__ == '__main__':
    root = tk.Tk()

    root.title("Masterlist Converter")
    # root.geometry("250x250")
    # root.resizable(0, 0)

    batchnum_label = tk.Label(text=u'Batch Number:')
    batchnum_entry = tk.Entry()

    pickup_date_label = tk.Label(text=u'Pickup Date:')
    pickup_date_entry = tk.Entry()

    download_button = tk.Button(text=u'Convert', command=convert)

    batchnum_label.grid(row=1, column=0)
    batchnum_entry.grid(row=1, column=1)

    pickup_date_label.grid(row=2, column=0)
    pickup_date_entry.grid(row=2, column=1)

    download_button.grid(row=3, column=1)

    root.mainloop()
