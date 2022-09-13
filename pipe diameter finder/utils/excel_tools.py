
# Need to filter out old rows that have been previously removed (typically at bottom of PSDs)
class last_row:
    last_psd_row = psd_data[psd_data['Full Data'] == "[END]"].index.to_list()[0] # Find [END] row/index 