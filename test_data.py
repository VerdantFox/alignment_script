test_wb_dict1 = {
    'THE': [15, 0], 'CAT': [18, 0], 'IN': [2, 0], 'HAT': [4, 0],
    'CCCDDDEE': [5, 95], 'FGHIKAA': [6, 94], 'OOPIS~JA': [7, 93],
    'JKA': [8, 92], 'BDEDF': [9, 91], 'SAT': [10, 0], 'ATOP': [11, 0],
    'VAT': [13, 0], 'OF': [14, 0], 'SKAT': [15, 0], 'HOW': [16, 0],
    'LIKE': [18, 0], 'CXCJKJ': [19, 81], 'ABCDEFG': [0, 100],
    'HIGJKLM': [0, 99], 'OPQRS~TU': [0, 98], 'VWXYZ': [0, 97],
    'AAAABBBB': [0, 96], 'EAJJKE': [0, 90], 'JKJLAJEL': [0, 89],
    'EUWOMJS': [0, 88], 'WJIS~JSO': [0, 87], 'SEUIS': [0, 86],
    'IEJUDU': [0, 85], 'SUIOUS': [0, 84], 'AJIFJF': [0, 83],
    'JDUIOS': [0, 82]
}

test_wb_dict2 = {
    'THE': [x for x in range(20)], 'CAT': [x*2 for x in range(20)],
    'IN': [x**2 for x in range(20)], 'HAT': [1 for x in range(20)],
    'CCCDDDEE': [x*2 % 3 for x in range(20)],
    'FGHIKAA': [(x+10) * 5 % 2 for x in range(20)],
    'OOPIS~JA': [(x+1) % 9 for x in range(20)],
}

print(test_wb_dict2)