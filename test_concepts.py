new_dict = dict()

new_dict['yellow'] = 3

if 'yellow' in new_dict:
    new_dict['yellow'] += 1
else:
    new_dict['yellow'] = 1

print(new_dict['yellow'])