import win32com.client as win32

# 色範囲の定義
orange_rgb_range = {
    'r': (245, 255),
    'g': (220, 235),
    'b': (205, 225)
}

fill_blue_rgb_range = {
    'r': (178, 188),
    'g': (207, 217),
    'b': (234, 244)
}

def hex_to_rgb(hex_color):
    if hex_color.startswith('#'):
        hex_color = hex_color[1:]
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    else:
        return (255, 255, 255)

def rgb_to_hex(rgb_color):
    return '#{:02x}{:02x}{:02x}'.format(rgb_color[0], rgb_color[1], rgb_color[2])

def is_rgb_in_range(rgb_color, color_range):
    return (color_range['r'][0] <= rgb_color[0] <= color_range['r'][1] and
            color_range['g'][0] <= rgb_color[1] <= color_range['g'][1] and
            color_range['b'][0] <= rgb_color[2] <= color_range['b'][1])

def get_cell_background_color(cell):
    shading = cell.Shading.BackgroundPatternColor
    rgb_color = (shading & 0xFF, (shading >> 8) & 0xFF, (shading >> 16) & 0xFF)
    
    if is_rgb_in_range(rgb_color, orange_rgb_range):
        return '#ffe8d1'
    elif is_rgb_in_range(rgb_color, fill_blue_rgb_range):
        return '#F0F8FF'
    
    return '#ffffff'

def create_html_table(table):
    html = '<table style="table-layout: fixed; width: 100%; text-align: center; border-collapse: collapse;">'
    html += '<tbody>'
    
    for i, row in enumerate(table.Rows):
        html += '<tr>'
        for j, cell in enumerate(row.Cells):
            cell_text = cell.Range.Text.strip()
            background_color = get_cell_background_color(cell)
            is_bold = cell.Range.Font.Bold
            
            if i == 0:
                if background_color == '#ffffff':
                    if is_bold:
                        html += f'<th><strong>{cell_text}</strong></th>'
                    else:
                        html += f'<th>{cell_text}</th>'
                else:
                    if is_bold:
                        html += f'<th style="background-color: {background_color};"><strong>{cell_text}</strong></th>'
                    else:
                        html += f'<th style="background-color: {background_color};">{cell_text}</th>'
            else:
                if j == 0:
                    if background_color == '#ffffff':
                        if is_bold:
                            html += f'<td><strong><a href="">{cell_text}</a></strong></td>'
                        else:
                            html += f'<td><a href="">{cell_text}</a></td>'
                    else:
                        if is_bold:
                            html += f'<td style="background-color: {background_color};"><strong><a href="">{cell_text}</a></strong></td>'
                        else:
                            html += f'<td style="background-color: {background_color};"><a href="">{cell_text}</a></td>'
                else:
                    if background_color == '#ffffff':
                        if is_bold:
                            html += f'<td><strong>{cell_text}</strong></td>'
                        else:
                            html += f'<td>{cell_text}</td>'
                    else:
                        if is_bold:
                            html += f'<td style="background-color: {background_color};"><strong>{cell_text}</strong></td>'
                        else:
                            html += f'<td style="background-color: {background_color};">{cell_text}</td>'
        html += '</tr>'
    
    html += '</tbody></table>'
    return html
