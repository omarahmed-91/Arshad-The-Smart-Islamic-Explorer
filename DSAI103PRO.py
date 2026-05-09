from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from serpapi import GoogleSearch
from tkinter import *
import customtkinter as ctk
import networkx as nx
import arabic_reshaper
from bidi.algorithm import get_display
from tkinter import messagebox
import pandas as pd
import openpyxl
import os
import matplotlib.pyplot as plt
from collections import Counter
import numpy as np

downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "نتائج_البحث.xlsx")

options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=options)

def save_to_excel(data, sheet_name="Sheet1"):
    df = pd.DataFrame(data)
    if os.path.exists(downloads_path):
        with pd.ExcelWriter(downloads_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(downloads_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def fix_arabic(text):
    reshaped_text = arabic_reshaper.reshape(text)
    bidi_text = get_display(reshaped_text)
    return bidi_text

def process_data_with_pandas(data_list):
    df = pd.DataFrame(data_list)
    if df.empty:
        return df
    df['title'] = df['title'].str.strip()
    df = df.drop_duplicates(subset=['title'])
    df = df[df['title'].str.len() > 2]
    return df

def islamweb_scrap(target_entry,target_reultbox):
    search_word=target_entry.get()
    driver.get(f"https://www.islamweb.net/ar/fatwa/?page=websearch&stxt={search_word}")
    time.sleep(2)
    results = driver.find_elements(By.CSS_SELECTOR, "h5 a")
    extracted_data = []
    for item in results:
        extracted_data.append({
            "title": item.text,
            "link": item.get_attribute("href")
        })
    df =pd.DataFrame(extracted_data)
    df['title'] =df['title'].str.strip()
    df_cleaned =df.drop_duplicates(subset=['title'])
    save_to_excel(df_cleaned, "الفتاوى")
    G.add_node(search_word,type="search")
    for index, row in df_cleaned.iterrows():
        clean_title = row['title']
        G.add_node(clean_title, type="fatwa")
        G.add_edge(search_word, clean_title)
    for i, row in df_cleaned.iterrows():
        driver.get(row['link'])
        time.sleep(1)
        footer_item = driver.find_element(By.CLASS_NAME, "footer-item")
        target_reultbox.insert('end',f"{i}- العنوان: {row['title']}\n","rtl")
        target_reultbox.insert('end',f"   الرابط: {row['link']}\n","rtl")
        target_reultbox.insert('end',f"   المعلومة الإضافية: {footer_item.text}\n","rtl")
        target_reultbox.insert('end',"-" * 20+"\n","rtl")

def youtube_serpapi_scrap(target_entry,target_reultbox,video_entry):

    search_word=target_entry.get()
    target_reultbox.insert('end',f"\n--- جاري البحث في يوتيوب عن: {search_word} ---\n","rtl")
    params = {
        "engine": "youtube",
        "search_query": search_word,
        "api_key": "906ecdea089f2b69f943ba09a0b7c3061658855317a9b2dca280acd5c8458252"
    }

    search = GoogleSearch(params)
    results = search.get_dict()
    videos = results.get("video_results", [])

    if not videos:
        target_reultbox.insert('end',"مفيش فيديوهات ظهرت في البحث.")
        return


    df = pd.DataFrame(videos)
    df['title'] = df['title'].str.strip().str.replace('\n', ' ', regex=False)
    df = df.drop_duplicates(subset=['title'])
    df['views'] = df['views'].fillna("غير معروف")
    df['channel_name'] = df['channel'].str.get('name').fillna("غير معروف")
    G.add_node(search_word, type="search")

    video_count = int(video_entry.get())

    youtube_data = []
    for i, row in df.head(video_count).iterrows():
        title = row['title']
        link = row.get('link', 'لا يوجد رابط')
        views = row['views']
        channel = row['channel_name']
        G.add_node(title, type="video")
        G.add_edge(search_word, title)
        target_reultbox.insert('end', f"- العنوان: {title}\n", "rtl")
        target_reultbox.insert('end', f"   الرابط: {link}\n", "rtl")
        target_reultbox.insert('end', f"   المشاهدات: {views}\n", "rtl")
        target_reultbox.insert('end', f"   القناة: {channel}\n", "rtl")
        target_reultbox.insert('end', "-" * 20 + "\n", "rtl")
    save_to_excel(youtube_data, "يوتيوب")

def get_books(target_entry,target_reultbox):

    search_word= target_entry.get()
    G.add_node(search_word, type="search")
    driver.get(f"https://ketabonline.com/ar/books?display=grid&q={search_word}&scope=titles&sort=_score")
    time.sleep(2)

    results = driver.find_elements(By.CLASS_NAME, "item-link-overlay")
    book_links = []
    for item in results:
        book_links.append(item.get_attribute("href"))

    raw_data_list=[]

    for link in book_links:
        driver.get(link)
        time.sleep(1)
        book_title = driver.find_element(By.CLASS_NAME, "strong7").text
        book_writer=driver.find_element(By.XPATH,'/html/body/div/div/div/div[2]/div[2]/section[1]/div/div[2]/div/div[2]/div/div[1]/div/div[2]/div/div[4]/div[2]').text
        date=driver.find_element(By.XPATH,"/html/body/div/div/div/div[2]/div[2]/section[1]/div/div[2]/div/div[2]/div/div[2]/div[2]/div[2]/ul/li[4]/span[2]").text
        download_book=driver.find_element(By.XPATH,'/html/body/div/div/div/div[2]/div[2]/section[1]/div/div[2]/div/div[2]/div/div[2]/div[2]/div[2]/ul/li[3]/span[2]').text
        page_book=driver.find_element(By.XPATH,'/html/body/div/div/div/div[2]/div[2]/section[1]/div/div[2]/div/div[2]/div/div[2]/div[2]/div[2]/ul/li[1]/span[2]').text

        if book_writer =='':
            writer_name = "غير معروف"
        else:
            writer_name = book_writer

        raw_data_list.append({
            "title": book_title,
            "writer":writer_name,
            "date": date,
            "download": download_book,
            "pages": page_book
        })

    df = pd.DataFrame(raw_data_list)
    df['title'] = df['title'].str.strip()
    df=df.drop_duplicates(subset=['title'])
    save_to_excel(df, "الكتب")
    G.add_node(search_word, type="search")
    for i, row in df.iterrows():
        G.add_node(row['title'], type="book")
        G.add_edge(search_word, row['title'])
        target_reultbox.insert('end', f" اسم الكتاب: {row['title']}\n", "rtl")
        target_reultbox.insert('end', f" أسم المؤلف: {row['writer']}\n", "rtl")
        target_reultbox.insert('end', f" {row['date']}\n", "rtl")
        target_reultbox.insert('end', f" {row['download']}\n", "rtl")
        target_reultbox.insert('end', f" {row['pages']}\n", "rtl")
        target_reultbox.insert('end', "-" * 20 + "\n", "rtl")

G = nx.Graph()
def show_graph():
    if len(G.nodes) == 0:
        messagebox.showinfo(title="تنبية",message="لا يوجد بيانات ")
        return

    plt.figure(figsize=(30, 25))
    pos = nx.spring_layout(G, k=1)

    node_colors = []
    for node, data in G.nodes(data=True):
        t = data.get("type")
        if t == "book":
            node_colors.append("green")
        elif t == "video":
            node_colors.append("red")
        elif t == "fatwa":
            node_colors.append("blue")
        else:
            node_colors.append("gray")

    labels = {}
    for node in G.nodes():
        short_text = node[:50]
        labels[node] = fix_arabic(short_text)

    nx.draw(G,pos,labels=labels,node_color=node_colors,node_size=400,font_size=7,edge_color="gray")
    plt.title("Knowledge Graph")
    plt.show()

def show_3d_point_cloud():
    if len(G.nodes) <= 1:
        messagebox.showinfo(title="تنبية", message="لا يوجد بيانات كافية.")
        return

    xs = []
    ys = []
    zs = []
    colors = []
    labels = []

    for node, data in G.nodes(data=True):
        t =data.get("type")
        val =len(node)

        if t == "video":
            xs.append(val)
            ys.append(0)
            zs.append(0)
            colors.append("red")
        elif t == "book":
            xs.append(0)
            ys.append(val)
            zs.append(0)
            colors.append("green")
        elif t == "fatwa":
            xs.append(0)
            ys.append(0)
            zs.append(val)
            colors.append("blue")
        else:
            continue
        labels.append(data.get("short_title", node[:15]))
    fig = plt.figure(figsize=(12, 10))
    ax = fig.add_subplot(projection='3d')
    ax.scatter(xs, ys, zs, c=colors, s=150, edgecolors='black', alpha=0.7)

    for i in range(len(xs)):
        ax.plot([0, xs[i]], [0, ys[i]], [0, zs[i]], color=colors[i], linestyle='--', alpha=0.3)
    ax.set_xlabel(fix_arabic("محور فيديوهات اليوتيوب (X)"))
    ax.set_ylabel(fix_arabic("محور الكتب والمؤلفات (Y)"))
    ax.set_zlabel(fix_arabic("محور الفتاوى الشرعية (Z)"))
    plt.title(fix_arabic("توزيع ثقل البيانات بين (فيديو - كتاب - فتوى)"))
    ax.view_init(elev=20, azim=45)

    plt.show()

def show_heatmap():
    if len(G.nodes) == 0:
        messagebox.showinfo(title="تنبية", message="لا يوجد بيانات كافية ")
        return
    all_titles = ""
    for node, data in G.nodes(data=True):
        if data.get("type") in ["book", "video", "fatwa"]:
            all_titles += " " + node

    words = [w for w in all_titles.split() if len(w) > 3]

    if not words:
        messagebox.showinfo(title="تنبية", message="العناوين قصيرة  للتحليل")
        return

    word_counts = Counter(words).most_common(10)
    labels = [fix_arabic(item[0]) for item in word_counts]
    values = [item[1] for item in word_counts]
    final_values = values[:10]
    while len(final_values) < 10: final_values.append(0)
    matrix = np.array(final_values).reshape(2, 5)
    label_matrix = np.array(labels + [""] * (10 - len(labels))).reshape(2, 5)
    fig, ax = plt.subplots(figsize=(12, 6))
    im = ax.imshow(matrix, cmap='YlGn')

    for i in range(2):
        for j in range(5):
            text = ax.text(j, i, f"{label_matrix[i, j]}\n({matrix[i, j]})",ha="center", va="center", color="black",fontsize=9, fontweight='bold')

    plt.colorbar(im, label=fix_arabic("تكرار الكلمة"))
    ax.set_xticks([])
    ax.set_yticks([])
    plt.title(fix_arabic("خريطة حرارية لأكثر الكلمات تكراراً في البحث"))
    plt.show()

windo=Tk()
windo.geometry("1100x650")
windo.title("أرشاد || الباحث الإسلامي الذكي")
icon= PhotoImage(file='pattern.png')
windo.iconphoto(True,icon)
windo.config(background="#0f172a")

def clear():
    for i in windo.winfo_children():
        i.destroy()

def main_menu():
    clear()
    image5 = PhotoImage(file='mandala.png')
    label2 = Label(windo, image=image5, background="#0f172a")
    label2.image = image5
    label2.pack(side='left')

    image6 = PhotoImage(file='mandala.png')
    label3 = Label(windo, image=image6, background="#0f172a")
    label3.image = image6
    label3.pack(side='right')
    label1=Label(windo,text="أختر مصدر البحث",font=('Times New Roman Baltic',25,'bold'),background='#0f172a',fg='white')
    label1.pack(pady=20)

    image1=PhotoImage(file='seo.png')
    btn_all=Button(windo,text="بحث كامل",
    font=('Times New Roman Baltic',13,'bold'),image=image1,compound="top",bg='#2563eb',fg='white',padx=40,width=120,
    command=all_search,relief=RAISED,bd=5)
    btn_all.image=image1
    btn_all.pack(pady=10)

    image2=PhotoImage(file='research.png')
    btn_for_book = Button(windo, text="بحث كتب فقط",
    font=('Times New Roman Baltic', 13, 'bold'), image=image2, compound="top", bg='#2563eb',
    fg='white', padx=40, width=120,
    command=search_for_book, relief=RAISED, bd=5)
    btn_for_book.image = image2
    btn_for_book.pack(pady=10)

    image3 = PhotoImage(file='find.png')
    btn_for_youtube = Button(windo, text="بحث دروس يوتيوب فقط",
    font=('Times New Roman Baltic', 13, 'bold'), image=image3, compound="top", bg='#2563eb',
    fg='white', padx=40, width=120,
    command=youtube_search, relief=RAISED, bd=5)
    btn_for_youtube.image = image3
    btn_for_youtube.pack(pady=10)

    image4 = PhotoImage(file='search.png')
    btn_for_youtube = Button(windo, text="بحث عن فتاوي فقط",
    font=('Times New Roman Baltic', 13, 'bold'), image=image4, compound="top", bg='#2563eb',
    fg='white', padx=40, width=120,
    command=foatwa_search, relief=RAISED, bd=5)
    btn_for_youtube.image = image4
    btn_for_youtube.pack(pady=10)

    image8 = PhotoImage(file='resized_info.png')
    btn_for_info = Button(windo, text="نبذة عن التطبيق",
    font=('Times New Roman Baltic', 12, 'bold'), image=image8, compound="top", bg='#2563eb',
    fg='white', padx=40, width=120,height=50,
    command=info_app, relief=RAISED, bd=5)
    btn_for_info.image = image8
    btn_for_info.pack(pady=10)

    btn_graph=Button(windo, text="networkx", font=('Times New Roman Baltic', 13, 'bold'), bg='#16a34a', fg='white',command=show_graph)
    btn_graph.pack(pady=10)

    btn_heatmap = Button(windo, text="heatmap", font=('Times New Roman Baltic', 13, 'bold'), bg='#16a34a', fg='white',command=show_heatmap)
    btn_heatmap.pack(pady=10)

    btn_3d = Button(windo, text="3d point", font=('Times New Roman Baltic', 13, 'bold'), bg='#16a34a', fg='white',command=show_3d_point_cloud)
    btn_3d.pack(pady=10)

def all_search():
    clear()
    global entry, video_count1, result_box
    entry=ctk.CTkEntry(windo,font=('Times New Roman Baltic',15),width=400,justify='right',placeholder_text="...مثال:الصلاة،الزكاة",height=35)
    entry.pack(pady=20)

    video_count1 = ctk.CTkEntry(windo, font=('Times New Roman Baltic', 15), width=400, justify='right',
    placeholder_text="عايز كام فديو..؟", height=35)
    video_count1.pack(pady=5)

    btn_sub = ctk.CTkButton(windo, text="أبحث",font=('Times New Roman Baltic',15,'bold'),
    command=run_full_search,corner_radius=10)
    btn_sub.pack(pady=10)

    btn_back = ctk.CTkButton(windo, text="رجوع الي القائمة",font=('Times New Roman Baltic',15,'bold'),
    command=main_menu,corner_radius=10)
    btn_back.pack(pady=10)

    result_box = ctk.CTkTextbox(windo, font=('Arial', 14), width=1000, height=500,
    corner_radius=15, border_width=2,border_color="#1f538d", text_color="white")
    result_box.tag_config("rtl", justify='right')
    result_box.pack(pady=20, padx=20)

def run_full_search():
    # result_box.insert('end', "\n...النتائج التي تم استخرجها من الأحاديث...\n", "rtl")
    # get_hadith_data(entry,result_box)

    result_box.insert('end', "\n\n...النتائج التي تم استخرجها من الكتب...\n\n", "rtl")
    get_books(entry,result_box)

    result_box.insert('end', "\n\n=============================================================\n\n", "rtl")
    result_box.insert('end', "\n\n...النتائج التي تم استخرجها من دروس اليوتيوب...\n\n", "rtl")
    youtube_serpapi_scrap(entry,result_box,video_count1)

    result_box.insert('end', "\n\n=============================================================\n\n", "rtl")
    result_box.insert('end', "\n\n... النتائج التي تم استخرجها من الفتاوي ...\n\n", "rtl")
    islamweb_scrap(entry,result_box)

def search_for_book():
    clear()
    global entry1, result_box1
    entry1 = ctk.CTkEntry(windo, font=('Times New Roman Baltic', 15), width=400, justify='right',
    placeholder_text="...مثال:الصلاة،الزكاة", height=35)
    entry1.pack(pady=20)

    btn_sub1 = ctk.CTkButton(windo, text="أبحث", font=('Times New Roman Baltic', 15, 'bold'),
    command=lambda:get_books(entry1,result_box1), corner_radius=10)
    btn_sub1.pack(pady=10)

    btn_back1= ctk.CTkButton(windo, text="رجوع الي القائمة", font=('Times New Roman Baltic', 15, 'bold'),
    command=main_menu, corner_radius=10)
    btn_back1.pack(pady=10)

    result_box1= ctk.CTkTextbox(windo, font=('Arial', 14), width=1000, height=500,
    corner_radius=15, border_width=2, border_color="#1f538d", text_color="white")
    result_box1.tag_config("rtl", justify='right')
    result_box1.pack(pady=20, padx=20)

def youtube_search():
    clear()
    global entry2, result_box2, video_count

    entry2= ctk.CTkEntry(windo, font=('Times New Roman Baltic', 15), width=400, justify='right',
    placeholder_text="...مثال:الصلاة،الزكاة", height=35)
    entry2.pack(pady=20)

    video_count =ctk.CTkEntry(windo, font=('Times New Roman Baltic', 15), width=400, justify='right',
    placeholder_text="عايز كام فديو..؟", height=35)
    video_count.pack(pady=5)

    btn_sub2 = ctk.CTkButton(windo, text="أبحث", font=('Times New Roman Baltic', 15, 'bold'),
    command=lambda:youtube_serpapi_scrap(entry2,result_box2,video_count), corner_radius=10)
    btn_sub2.pack(pady=10)

    btn_back2 = ctk.CTkButton(windo, text="رجوع الي القائمة", font=('Times New Roman Baltic', 15, 'bold'),
    command=main_menu, corner_radius=10)
    btn_back2.pack(pady=10)

    result_box2 = ctk.CTkTextbox(windo, font=('Arial', 14), width=1000, height=500,
    corner_radius=15, border_width=2, border_color="#1f538d", text_color="white")
    result_box2.tag_config("rtl", justify='right')
    result_box2.pack(pady=20, padx=20)

def foatwa_search():
    clear()
    global entry3, result_box3
    entry3 = ctk.CTkEntry(windo, font=('Times New Roman Baltic', 15), width=400, justify='right',
    placeholder_text="...مثال:الصلاة،الزكاة", height=35)
    entry3.pack(pady=20)

    btn_sub3 = ctk.CTkButton(windo, text="أبحث", font=('Times New Roman Baltic', 15, 'bold'),
    command=lambda: islamweb_scrap(entry3, result_box3), corner_radius=10)
    btn_sub3.pack(pady=10)

    btn_back3 = ctk.CTkButton(windo, text="رجوع الي القائمة", font=('Times New Roman Baltic', 15, 'bold'),
    command=main_menu, corner_radius=10)
    btn_back3.pack(pady=10)

    result_box3 = ctk.CTkTextbox(windo, font=('Arial', 14), width=1000, height=500,
    corner_radius=15, border_width=2, border_color="#1f538d", text_color="white")
    result_box3.tag_config("rtl", justify='right')
    result_box3.pack(pady=20, padx=20)

def info_app():
    clear()
    label4 = Label(windo, text="هذا التطبيق يهدف الي توحيد مصدر واحد للبحث\nوتطوير نظام البحث الاليكتروني للمستخدمين\nالمطورون:\nعمر أحمد و عمرو أشرف\n النسخه:v.1.0.0",
    font=('Times New Roman Baltic', 30, 'bold'), background='#0f172a',
    fg='white')
    label4.pack(pady=50)

    btn_back4 = ctk.CTkButton(windo, text="رجوع الي القائمة", font=('Times New Roman Baltic', 15, 'bold'),command=main_menu, corner_radius=10)
    btn_back4.pack(pady=30)

main_menu()
windo.mainloop()