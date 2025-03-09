import pandas as pd
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from collections import defaultdict


def load_data(file_path):
    df = pd.read_excel(file_path)
    df = df.where(pd.notna(df), None)
    return df


def organize_wine_data(df):
    wine_dict = defaultdict(list)
    for _, row in df.iterrows():
        category = row["Категория"]
        wine_info = {
            "Название": row["Название"],
            "Сорт": row["Сорт"],
            "Цена": row["Цена"],
            "Картинка": row["Картинка"],
            "Акция": row["Акция"],
        }
        wine_dict[category].append(wine_info)
    return wine_dict


def get_year_word(number):
    if 11 <= number % 100 <= 19:
        return "лет"
    last_digit = number % 10
    if last_digit == 1:
        return "год"
    if 2 <= last_digit <= 4:
        return "года"
    return "лет"


def render_html(wine_dict, years):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    formatted_years = f"{years} {get_year_word(years)}"
    
    return template.render(
        datetime=formatted_years,
        wine_categories=dict(wine_dict)
    )


def save_html(content, filename='index.html'):
    with open(filename, 'w', encoding="utf8") as file:
        file.write(content)


def run_server():
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    file_path = "wine3.xlsx"
    df = load_data(file_path)
    wine_dict = organize_wine_data(df)
    rendered_page = render_html(wine_dict, 102)
    save_html(rendered_page)
    run_server()