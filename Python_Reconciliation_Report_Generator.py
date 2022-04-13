import pandas as pd, numpy as np

# let's load the csv data to pandas DataFrame
# Company X's (cx) Data is in three files
cx_sku_weight = pd.read_excel(r"/content/drive/MyDrive/Colab Notebooks/Solve_Problem/Company X - SKU Master.xlsx")
cx_order = pd.read_excel(r"/content/drive/MyDrive/Colab Notebooks/Solve_Problem/Company X - Order Report.xlsx")
cx_pincode_zone = pd.read_excel(r"/content/drive/MyDrive/Colab Notebooks/Solve_Problem/Company X - Pincode Zones.xlsx")

# Courier Companies (cc) data is in two files
cc_rates = pd.read_excel(r"/content/drive/MyDrive/Colab Notebooks/Solve_Problem/Courier Company - Rates.xlsx")
cc_invoice = pd.read_excel(r"/content/drive/MyDrive/Colab Notebooks/Solve_Problem/Courier Company - Invoice.xlsx")

# self defined weight calculator to c
def weight_slab (n):
    """ This function converts grams to weights in kg and rounds up to nearest 0.5 multiple.
        Logic: We find a value to add to n such that it rounds up to nearest 0.5 multiple.
        eg: - If the total weight is 400 gram then weight slab should be 0.5
            - If the total weight is 950 gram then weight slab should be 1
            - If the total weight is 1 KG then weight slab should be 1
            - If the total weight is 2.2 KG then weight slab should be 2.5
    """
    n = n/1000  # Grm to Kgs conversion
    return (n + (0.5 - n) % 0.5)

# self defined billing calculator
def rate(target):
    return cc_rates[target][0]


def billing_calc(weight, zone, Type_of_Shipment):

    if Type_of_Shipment == 'Forward charges':
        shipment_type_fwd = 'fwd'
        fixed_target = "_".join([shipment_type_fwd, zone, 'fixed'])
        additional_target = "_".join([shipment_type_fwd, zone, 'additional'])

        fwd_fxd_rate = rate(fixed_target)
        fwd_add_rate = rate(additional_target)

        fwd_fxd_chrges = fwd_fxd_rate  # 1

        fwd_add_chrges = add_wght = weight - 0.5  # additional weight
        add_multiples_of_half_kg = 0  # consider additional multiple to be zero
        while add_wght > 0:
            add_wght = add_wght - 0.5
            add_multiples_of_half_kg += 1  # no of multiples of 0.5 KG
        fwd_add_chrges = fwd_add_rate * add_multiples_of_half_kg  # 2

        rto_fxd_chrges = 0  # 3
        rto_add_chrges = 0  # 4

    else:
        shipment_type_fwd = 'fwd'
        shipment_type_rto = 'rto'
        fwd_fixed_target = "_".join([shipment_type_fwd, zone, 'fixed'])
        fwd_additional_target = "_".join([shipment_type_fwd, zone, 'additional'])

        rto_fixed_target = "_".join([shipment_type_rto, zone, 'fixed'])
        rto_additional_target = "_".join([shipment_type_rto, zone, 'additional'])

        fwd_fxd_rate = rate(fwd_fixed_target)
        fwd_add_rate = rate(fwd_additional_target)
        rto_fxd_rate = rate(rto_fixed_target)
        rto_add_rate = rate(rto_additional_target)

        fwd_fxd_chrges = fwd_fxd_rate  # 1

        add_wght = weight - 0.5  # additional weight
        add_multiples_of_half_kg = 0  # consider additional multiple to be zero
        while add_wght > 0:
            add_wght = add_wght - 0.5
            add_multiples_of_half_kg += 1  # no of multiples of 0.5 KG
        fwd_add_chrges = fwd_add_rate * add_multiples_of_half_kg  # 2

        rto_fxd_chrges = rto_fxd_rate  # 3

        add_wght1 = weight - 0.5  # additional weight
        add_multiples_of_half_kg_1 = 0  # consider additional multiple to be zero
        while add_wght1 > 0:
            add_wght1 = add_wght1 - 0.5
            add_multiples_of_half_kg_1 += 1  # no of multiples of 0.5 KG
        rto_add_chrges = rto_add_rate * add_multiples_of_half_kg_1  # 4

    total = round(fwd_fxd_chrges + fwd_add_chrges + rto_fxd_chrges + rto_add_chrges, 3)  # 1+2+3+4

    return total


def main():
    """
    Runs the main Programe code
    :return: if successfull returns successfully reconciliation report generated.
    """
    # We create a new dataframe order_report
    order_report = cx_order.copy()  # create copy of Company X order report for further processing

    pd.merge(order_report, cx_sku_weight, on="SKU") # We check the merged dataframe on common values columns "SKU"


    order_report = pd.merge(order_report, cx_sku_weight, on="SKU") # creating a new DF after merging on SKU column
    order_report.rename(columns={"Weight (g)": "Weight_(g)_per_unit"},inplace= True) # renaming
    order_report.rename(columns={"ExternOrderNo": "Order ID"},inplace= True)

    order_report['Weight_per_SKU_per_order'] = order_report['Weight_(g)_per_unit'] * order_report["Order Qty"]

    cum_weight_per_order = pd.DataFrame(order_report.groupby(by = ["Order ID"])["Weight_per_SKU_per_order"].sum())
    cum_weight_per_order.rename(columns={"Weight_per_SKU_per_order": "Total weight as per X (G)"},inplace= True)
    cum_weight_per_order.rename(columns={"ExternOrderNo": "Order ID"},inplace =True)
    cum_weight_per_order # cumulative sum of weights in one order


    order_report = pd.merge(order_report, cum_weight_per_order, on="Order ID") # merge total weight per order in Kg to order report


    cum_weight_per_order["Weight slab as per X (KG)"] = weight_slab(cum_weight_per_order["Total weight as per X (G)"])

    # creating a new columns with values in adjusted weights in KG
    order_report["Weight slab as per X (KG)"] = weight_slab(order_report["Total weight as per X (G)"])

    # Expected output df
    Result_df = pd.concat([cc_invoice["Order ID"],cc_invoice['AWB Code']],axis=1)


    Result_df = pd.merge(Result_df,cum_weight_per_order, on = 'Order ID')
    Result_df["Total weight as per X (G)"] = Result_df["Total weight as per X (G)"]/1000
    Result_df.rename(columns= {"Total weight as per X (G)": "Total weight as per X (KG)"},inplace = True)


    # As we want to add CC charged weight column to Result_df, but indexing is different.
    cc_invoice_OrderID_indx = cc_invoice.copy() # created new copy df of courier company invoice copy,
    cc_invoice_OrderID_indx.set_index("Order ID",inplace= True) # set index as Order ID

    Result_df = pd.merge(Result_df,cc_invoice_OrderID_indx["Charged Weight"], on= "Order ID")
    Result_df.rename(columns={"Charged Weight": "Total weight as per Courier Company (KG)"},inplace=True)

    Result_df["Weight slab charged by Courier Company (KG)"] = weight_slab(Result_df["Total weight as per Courier Company (KG)"]*1000) #using our function weights slab which takes values in grms

    cx_pincode_zone.set_index(cc_invoice_OrderID_indx.index,inplace= True) # We assign index to cx_pincode_zone


    Result_df = pd.merge(Result_df,cx_pincode_zone["Zone"], on = 'Order ID') #Delivery Zone as per X
    Result_df.rename(columns={'Zone':'Delivery Zone as per X'},inplace = True)


    Result_df = pd.merge(Result_df,cc_invoice_OrderID_indx["Zone"], on = 'Order ID') # Delivery Zone charged by Courier Company
    Result_df.rename(columns={'Zone':'Delivery Zone charged by Courier Company'},inplace = True)

    # create a new columns just for replacing it with new values with same rows
    Result_df['Expected Charge as per X (Rs.)']= Result_df['Order ID']

    for i in Result_df.index:
        Result_df['Expected Charge as per X (Rs.)'][i] = billing_calc(Result_df['Weight slab as per X (KG)'][i],
                                                                      Result_df['Delivery Zone as per X'][i],
                                                                      cc_invoice['Type of Shipment'][i])
        #df['Charges Billed by Courier Company (Rs.)'][i] = billing_calc(df['Order ID'][i],df['Weight slab charged by Courier Company (KG)'][i],
        # df['Delivery Zone charged by Courier Company'][i],cc_invoice_OrderID_indx['Type of Shipment'][i])


    # billing amount from cc_invoice to be added to Result_df
    Result_df = pd.merge(Result_df, cc_invoice_OrderID_indx['Billing Amount (Rs.)'], on =['Order ID'])
    Result_df.rename(columns={'Billing Amount (Rs.)':'Charges Billed by Courier Company (Rs.)'},inplace= True)


    Result_df['Difference Between Expected Charges and Billed Charges (Rs.)'] =  Result_df['Expected Charge as per X (Rs.)'] - \
                                                                                 Result_df['Charges Billed by Courier Company (Rs.)']
    #Output 2

    #Correctly Charged
    Count1 = Result_df[Result_df['Difference Between Expected Charges and Billed Charges (Rs.)']==0].shape[0]

    #OverCharged
    Count2 = Result_df[Result_df['Difference Between Expected Charges and Billed Charges (Rs.)']<0].shape[0]

    # UnderCharged
    Count3 = Result_df[Result_df['Difference Between Expected Charges and Billed Charges (Rs.)']>0].shape[0]

    # Correctly Charged
    Amount1 = Result_df['Charges Billed by Courier Company (Rs.)'][Result_df['Difference Between Expected Charges and Billed Charges (Rs.)']==0].sum()

    # Overcharged
    Amount2 = Result_df['Charges Billed by Courier Company (Rs.)'][Result_df['Difference Between Expected Charges and Billed Charges (Rs.)']<0].sum()

    # UnderCharged
    Amount3 = Result_df['Charges Billed by Courier Company (Rs.)'][Result_df['Difference Between Expected Charges and Billed Charges (Rs.)']>0].sum()


    summary_df = pd.DataFrame ({'Count': [Count1,Count2,Count3],'Amount':[Amount1,Amount2, Amount3]},index= ['Total Orders - Correctly Charged','Total Orders - Over Charged','Total Orders - Under Charged'])



    with pd.ExcelWriter('Reconciliation_Report.xlsx') as writer:
        summary_df.to_excel(writer, sheet_name='Summary')
        Result_df.to_excel(writer, sheet_name='Calculations', index = False)

    return print("Reconciliation Report Successfully Generated!")



if __name__ == '__main__':
    main()