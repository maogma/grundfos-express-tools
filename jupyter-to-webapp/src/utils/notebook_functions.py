def add_curve_data_points(parent_elem, curve_type, df):
    """ Adds curve data points for flow/power, flow/head, flow/NPSH """

    for _, row in df.iterrows():
        datapoint_elem = ET.SubElement(parent_elem, "DataPoint", disabled="false")
        
        if curve_type == 'Efficiency':
            datapoint_dict = {
                'x': row['Q [m³/h]'],
                'y': row['Eta1'],
                'isOnCurve': 'false',
                'division': 'false',
                'useCubicSplines': 'false',
                'slopeEnabled': 'false'
            }
        
        elif curve_type == 'Power':
            datapoint_dict = {
                'x': row['Q [m³/h]'],
                'y': row['P2 [kW]'],
                'isOnCurve': 'false',
                'division': 'false',
                'useCubicSplines': 'false',
                'slopeEnabled': 'false'
            }

        elif curve_type == 'Head':
            datapoint_dict = {
                'x': row['Q [m³/h]'],
                'y': row['H [m]'],
                'isOnCurve': 'false',
                'division': 'false',
                'useCubicSplines': 'true',
                'slopeEnabled': 'false'
            }

        elif curve_type == 'NPSH':
            datapoint_dict = {
                'x': row['Q [m³/h]'],
                'y': row['NPSH [m]'],
                'isOnCurve': 'false',
                'division': 'false',
                'useCubicSplines': 'false',
                'slopeEnabled': 'false'
            }

        else:
            print(f'curve_type not allowed: {curve_type}')
            
        add_elem_from_dict(datapoint_elem, datapoint_dict)