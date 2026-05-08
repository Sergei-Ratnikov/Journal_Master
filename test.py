import re


def parse_route_to_bounds(route_string):
    '''
    преобразование кординат размерной сетки генплана типа 16E; 14N в координаты квадрата x, y
    Args:
        route_string: строка, которая может содержать координаты типа 16E; 14N
    '''
    pattern = re.compile(r'(\d{1,2})\s*([EN])')
    matches = pattern.findall(route_string.upper())
    print(matches)
    if len(matches) < 2:
        return {'is_valid': False, 'original': route_string}
    x_val = None
    y_val = None
    for num, axis in matches:
        print('qqqqq')
        print(num)
        print(axis)
    #     value = int(num) * 100
    #     if axis == 'E':
    #         y_val = value
    #     elif axis == 'N':
    #         x_val = value
    # if x_val is None or y_val is None:
    #     return {'is_valid': False, 'original': route_string}
    # return {
    #     'is_valid': True,
    #     'x_min': x_val,
    #     'x_max': x_val + 100,
    #     'y_min': y_val,
    #     'y_max': y_val + 100
    # }

parse_route_to_bounds('16E; 14-30N;')