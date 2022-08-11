


# Need to filter out old rows that have been previously removed (typically at bottom of PSDs)
last_psd_row = psd_data[psd_data['Full Data'] == "[END]"].index.to_list()[0] # Find [END] row/index 