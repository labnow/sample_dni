# import json

# with open('all_dict.json', 'r') as j:
#     json_data = json.load(j)

# json_data['data']['1'] = [1,2,3,4,5]

# with open('all_dict.json', 'w') as json_file:
#     json.dump(json_data, json_file, ensure_ascii=False, indent=4)
import re
a = 'abc-asd.asd'
b = re.sub(r'[-.]', '', a)
print(b)