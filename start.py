import pandas as pd
import jinja2
from datetime import datetime
import webbrowser
import os
import requests

def load_customer_data(file_path):
    return pd.read_excel(file_path)

def calculate_trip_days(start_date, end_date):
    start_date = datetime.strptime(start_date, "%Y-%m-%d")
    end_date = datetime.strptime(end_date, "%Y-%m-%d")
    return (end_date - start_date).days + 1

def create_google_maps_embed_url(api_key, origin, destination):
    base_url = "https://www.google.com/maps/embed/v1/directions"
    params = f"?key={api_key}&origin={origin}&destination={destination}&avoid=tolls|highways"
    return base_url + params

def generate_html(plan, maps_urls, filename='travel_plan.html'):
    template_loader = jinja2.FileSystemLoader(searchpath='./')
    template_env = jinja2.Environment(loader=template_loader)
    template_file = 'template.html'
    template = template_env.get_template(template_file)

    html_output = template.render(plan=plan, maps_urls=maps_urls)
    with open(filename, 'w',encoding="utf-8") as file:
        file.write(html_output)

def get_travel_time_and_distance(api_key, origin, destination):
    """
    获取每段行程的旅行时间和距离

    Args:
        api_key: Google Maps API 密钥
        origin: 起点
        destination: 终点

    Returns:
        旅行时间 (单位：秒)
        距离 (单位：米)
    """

    url = "https://maps.googleapis.com/maps/api/directions/json?origin={}&destination={}&key={}".format(
        origin, destination, api_key
    )
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()

    if data.get("routes"):
        travel_time = data["routes"][0]["legs"][0]["duration"]["value"] / 3600
        distance = data["routes"][0]["legs"][0]["distance"]["value"] / 1000
    else:
        travel_time = None
        distance = None

    return travel_time, distance

def travel_plan_program():
    customer_data = load_customer_data('clients.xlsx')
    start_date = input("输入出差开始日期 (格式：YYYY-MM-DD): ")
    end_date = input("输入出差结束日期 (格式：YYYY-MM-DD): ")
    api_key = "" 
    customers_per_day = int(input("输入每天计划拜访的客户数量: "))
    total_days = calculate_trip_days(start_date, end_date)

    plan = {}
    maps_urls = {}
    last_destination = None  # 存储上一天的终点

    for day in range(total_days):
        if last_destination is not None:
            # 将上一天的终点作为今天的起点
            day_customers = customer_data.sample(customers_per_day - 1)
            day_customers.loc[-1] = {'客户地址': last_destination}  # 添加起点
            day_customers.index = day_customers.index + 1
            day_customers = day_customers.sort_index()
        else:
            day_customers = customer_data.sample(customers_per_day)

        plan[day + 1] = day_customers
        origin = day_customers.iloc[0]['客户地址']
        destination = day_customers.iloc[-1]['客户地址']
        last_destination = destination  # 更新终点为下一天的起点

        maps_urls[day + 1] = create_google_maps_embed_url(api_key, origin, destination)
        # TODO: 在这里调用Google Maps API获取每段行程的旅行时间和距离
        travel_time, distance = get_travel_time_and_distance(api_key, origin, destination)
        plan[day + 1]["旅行时间"] = travel_time
        plan[day + 1]["距离"] = distance

    generate_html(plan, maps_urls)
    # TODO: 根据旅行时间、距离和费用等信息生成多个行程规划方案

travel_plan_program()

filename = 'travel_plan.html'
file_path = os.path.abspath(filename)
webbrowser.open(file_path)
