import json
from to_excel import *
from bs4 import BeautifulSoup

def html_to_text(html):
    soup = BeautifulSoup(html, "html.parser")

    for br in soup.select("br"):
        br.replace_with("\n")

    for p in soup.find_all('p'):
        p.replace_with(p.text + ' \n')
    
    return soup.get_text()


with open("Медиаполе.json", 'r', encoding='utf-8') as f:
    boards_i_tasks = json.load(f)

stick_dick = {}
for sticker in boards_i_tasks['stickers']:
    stick_dick[sticker['id']] = {}
    stick_dick[sticker['id']]['title'] = sticker['title']
    for state_id, state_info in sticker['states']['index'].items():
        stick_dick[sticker['id']][state_id] = state_info['name']

#id некорректных stickers
stick_dick['4fce20d2-24d0-4be0-b0df-0736750da030'] = "Количество карт"
stick_dick['0fb4e2ac-df62-43f5-b381-11674465f8dc'] = "Поле"
stick_dick['5d0ec892-d575-47ed-8ec0-70ebc0f6c006'] = "Стоимость"

columns_dick = {}
for board in boards_i_tasks['boards']:
    for columns in board["columns"]:
        for i in columns['tasks']:
            columns_dick[i] = columns['title']

tasks_dick = {}
for task_id, task_info in boards_i_tasks['tasks'].items():
    tasks_dick[task_id] = {}
    tasks_dick[task_id]['title'] = task_info['title']
    tasks_dick[task_id]['column'] = columns_dick[task_id]
    tasks_dick[task_id]['description'] = html_to_text(task_info['description'])
    tasks_dick[task_id]['stickers'] = []
    for stick_id, stick_state in task_info['stickers'].items():
        try:
            if stick_state.replace(' ', '').isdigit() or stick_state.replace(' ', '').split(',')[0].isdigit():
                state_title = stick_dick[stick_id]+': '+stick_state
                tasks_dick[task_id]['stickers'].append(state_title)
            else:
                state_title = stick_dick[stick_id]['title']+': '+stick_dick[stick_id][stick_state]
                tasks_dick[task_id]['stickers'].append(state_title)
        except KeyError as e:

            print(e)

if __name__ == "__main__":
    create_excel(tasks_dick)