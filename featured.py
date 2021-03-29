featured_products_list = ['CAJ 4151', 
					'CAK 7714', 
					'CAK 61642', 
					'CAK 61689', 
					'CAK 61945', 
					'CAK 62581', 
					'CWK 62287', 
					'CWJ 9181']

def featured_products(ws):
	ws.insert_cols(22)
	ws['V1'].value = 'Featured'
	index_counter = 1
	for sku in ws['A']:
		sku = sku.value
		if sku in featured_products_list:
			ws[f'V{index_counter}'].value = 'yes'
			index_counter += 1
		else:
			index_counter += 1