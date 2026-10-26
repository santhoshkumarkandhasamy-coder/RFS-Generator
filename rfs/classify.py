def classify(p, g, n):
    nl=n.lower(); gl=g.lower()
    if p=='Profit & Loss A/c': return ('XX_PLA','XX_PLA')
    if p=='Branch / Divisions': return ('XX_BR','XX_BR')
    if n=='Opening Stock': return ('XX_OPSTK','XX_OPSTK')
    if p=='Capital Account':
        return ('RS','RS') if 'reserve' in nl else ('SC','SC')
    if p=='Loans (Liability)':
        if g=='Bank OD A/c': return ('BANK','STB_BANK')
        if g=='Director Loan': return ('STLA_OTH','LTB_UNS')
        if g=='Long Term Borrowing': return ('STLA_OTH','LTB_SEC')
        if g=='Credit Cards' or 'credit card' in nl: return ('STLA_OTH','OCL_EXP')
        return ('STLA_OTH','STB_OTH')
    if p=='Current Liabilities':
        if gl.startswith('sundry creditors - msme'): return ('STLA_SUP','TP_MSME')
        if gl.startswith('sundry creditors'): return ('STLA_SUP','TP_OTH')
        if g=='Long Term Provisions': return ('STLA_OTH','LTP_GRAT')
        if g in ('Other Long Term Liabilities','Rental Deposit Payable'): return ('STLA_OTH','OLTL')
        if g=='Short Term Provisions': return ('OCA_TAX','STP_TAX')
        if g=='Credit Cards': return ('STLA_OTH','OCL_EXP')
        if g=='Salary Provisions':
            if 'audit fee' in nl or 'rent payable' in nl: return ('STLA_OTH','OCL_EXP')
            if 'pf payable' in nl or 'esi payable' in nl or 'professional tax' in nl: return ('OCA_GOVT','OCL_STAT')
            return ('STLA_EMP','OCL_EMP')
        if g in ('Other Current Liabilities','Provisions'): return ('STLA_OTH','OCL_EXP')
        if 'tds recoverable' in nl or 'tds receivable' in nl or 'income tax refund' in nl: return ('OCA_TAX','OCL_STAT')
        return ('OCA_GOVT','OCL_STAT')
    if p=='Fixed Assets':
        if 'building work in progress' in nl: return ('CWIP','CWIP')
        if 'land' in nl and 'building' not in nl: return ('PPE_LAND','PPE_LAND')
        if 'building' in nl: return ('PPE_BLDG','PPE_BLDG')
        if g in ('Intangible Asset','Intangible Asset Under Development'): return ('IAUD','IAUD')
        if g=='Computer Software': return ('INT_SW','INT_SW')
        if g=='Computers': return ('PPE_COMP','PPE_COMP')
        if g=='Furniture & Fittings': return ('PPE_FURN','PPE_FURN')
        if g in ('Vehicles','Bike'): return ('PPE_VEH','PPE_VEH')
        if g in ('Electrical Fittings','Electrical and Fittings'): return ('PPE_ELEC','PPE_ELEC')
        return ('PPE_OFFEQ','PPE_OFFEQ')
    if p=='Current Assets':
        if g=='Deferred Tax Assets Net': return ('DTA','DTL')
        if g=='Sundry Debtors': return ('TR','OCL_ADV')
        if g in ('Long Term Loans and Advances','Rental Deposit Receivable') or 'security deposit' in nl or 'deposit - indian' in nl: return ('LTLA','OCL_EXP')
        if g=='Cash-in-hand': return ('CASH','OCL_EXP')
        if g=='Bank Accounts': return ('BANK','STB_BANK')
        if 'fixed deposit' in gl or 'fixed deposit' in nl or nl.startswith('fd -') or 'bank guarantee' in nl: return ('BANK_FD','OCL_EXP')
        if g=='IPO - Advances': return ('OCA_IPO','OCL_EXP')
        if g in ('GST Cash Ledger','GST Credit Ledger'): return ('OCA_GOVT','OCL_STAT')
        if 'unbilled' in nl: return ('OCA_UNB','OCL_EXP')
        if 'prepaid' in nl or 'prepaid expenses' in gl: return ('OCA_PRE','OCL_EXP')
        if 'emd' in nl: return ('OCA_EMD','OCL_EXP')
        if 'interest receivable' in nl or 'accrued interest' in nl: return ('OCA_INT','OCL_EXP')
        if 'salary advance' in nl or 'staff' in nl: return ('STLA_EMP','OCL_EMP')
        if 'tds' in nl or 'tcs' in nl or g=='Prepaid Taxes' or 'tcs/tds' in gl: return ('OCA_TAX','OCL_STAT')
        if g in ('Short Term Loans and Advances','Loans & Advances (Asset)'): return ('STLA_OTH','OCL_EXP')
        return ('OCA_OTH','OCL_EXP')
    if p=='Sales Accounts':
        if 'freight' in nl or 'frieght' in nl: c='REV_OTH'
        elif '(sos)' in nl or 'services' in nl or g=='AMC' or 'unbilled' in nl: c='REV_SERV'
        else: c='REV_PROD'
        return (c,c)
    if p=='Purchase Accounts': return ('PUR','PUR')
    if p=='Direct Expense':
        c='OE_FRT' if g=='Freight Charges' else 'DIR'; return (c,c)
    if p=='Indirect Incomes':
        c='OI_INT' if 'interest' in nl else 'OI_OTH'; return (c,c)
    if p=='Indirect Expenses':
        c=None
        if n=='Depreciation': c='DEP'
        elif g=='Tax Expenses': c='TAX_DEF' if 'deferred' in nl else 'TAX_CUR'
        elif g=='Financial Charges': c='FIN_OTH' if any(k in nl for k in ['bank charges','processing','credit card']) else 'FIN_INT'
        elif 'commitment fee' in nl: c='FIN_OTH'
        elif 'director remuneration' in nl: c='EMP_DIR'
        elif g in ('Employee Benefit Expenses','Salary A/c','ESI & PF Contribution','Staff Welfare'):
            if any(k in nl for k in ['pf ','pf employer','esi','admin charges','labour welfare']): c='EMP_PF'
            elif 'gratuity' in nl: c='EMP_GRAT'
            elif any(k in nl for k in ['welfare','food','education','accommodation','insurance']): c='EMP_WEL'
            else: c='EMP_SAL'
        elif 'health insurance' in nl or 'group insurance' in nl or 'employees group' in nl: c='EMP_WEL'
        elif g in ('Insurance Charges','Insurance Expenses'): c='OE_INS'
        elif 'audit fee' in nl: c='OE_AUD'
        elif 'professional' in nl: c='OE_PROF'
        elif 'electricity' in nl: c='OE_POWER'
        elif g=='Rental Charges': c='OE_RENT'
        elif g=='Rates & Taxes': c='OE_MISC' if 'round off' in nl else 'OE_RATES'
        elif 'ineligible' in nl: c='OE_RATES'
        elif 'repairs' in nl or g=='Vehicle Maintenance': c='OE_RM'
        elif g in ('Travelling & Conveyance','Travelling Allowance'): c='OE_TRAV'
        elif 'brokerage' in nl or 'commission' in nl: c='OE_COMM'
        elif 'advertisement' in nl or 'business promotion' in nl: c='OE_ADV'
        elif g in ('Telephone & Internet','Postage & Courier Charges'): c='OE_TEL'
        elif 'printing' in nl: c='OE_PRINT'
        elif g=='Freight Charges' or 'freight' in nl or 'transportation' in nl or 'unloading' in nl: c='OE_FRT'
        elif 'bad debts' in nl or 'written off' in nl: c='OE_BAD'
        elif 'forex' in nl: c='OE_FOREX'
        else: c='OE_MISC'
        return (c,c)
    return ('UNMAPPED','UNMAPPED')
