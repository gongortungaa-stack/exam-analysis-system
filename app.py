from flask import Flask, render_template, request, send_file
import pandas as pd
import matplotlib.pyplot as plt
import io
import base64

app = Flask(__name__)
df_global = None

@app.route('/')
def home():
    return render_template("index.html")

@app.route('/upload', methods=['POST'])
def upload():
    global df_global

    file = request.files['file']
    df = pd.read_excel(file)

    question_cols = ['д1', 'д2', 'д3', 'д4', 'д5', 'д6']
    df['нийт оноо'] = df[question_cols].sum(axis=1)

    average = round(df['нийт оноо'].mean(), 2)
    max_score = df['нийт оноо'].max()
    min_score = df['нийт оноо'].min()

    difficulty = df[question_cols].mean()
    hardest = difficulty.idxmin()
    easiest = difficulty.idxmax()

    plt.figure()
    difficulty.plot(kind='bar')
    plt.title("Даалгаврын гүйцэтгэл")

    img = io.BytesIO()
    plt.savefig(img, format='png')
    img.seek(0)
    graph_url = base64.b64encode(img.getvalue()).decode()
    plt.close()

    df_global = df

    return render_template(
        "result.html",
        tables=df.to_html(index=False),
        average=average,
        max_score=max_score,
        min_score=min_score,
        hardest=hardest,
        easiest=easiest,
        graph=graph_url
    )

@app.route('/download')
def download():
    global df_global

    if df_global is None:
        return "Эхлээд Excel файл upload хийнэ үү."

    question_cols = ['д1', 'д2', 'д3', 'д4', 'д5', 'д6']

    average = round(df_global['нийт оноо'].mean(), 2)
    max_score = df_global['нийт оноо'].max()
    min_score = df_global['нийт оноо'].min()

    difficulty = df_global[question_cols].mean()
    hardest = difficulty.idxmin()
    easiest = difficulty.idxmax()

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1-р sheet: Дүнгийн хүснэгт
        df_global.to_excel(writer, sheet_name='Дүн', index=False)

        # 2-р sheet: Анализ
        summary = pd.DataFrame({
            'Үзүүлэлт': [
                'Дундаж оноо',
                'Хамгийн их оноо',
                'Хамгийн бага оноо',
                'Хамгийн хэцүү даалгавар',
                'Хамгийн амар даалгавар'
            ],
            'Утга': [
                average,
                max_score,
                min_score,
                hardest,
                easiest
            ]
        })

        summary.to_excel(writer, sheet_name='Анализ', index=False)

        # 3-р sheet: Даалгаврын гүйцэтгэл
        task_result = difficulty.reset_index()
        task_result.columns = ['Даалгавар', 'Гүйцэтгэлийн хувь']
        task_result.to_excel(writer, sheet_name='Даалгавар', index=False)

    output.seek(0)

    return send_file(
        output,
        download_name="exam_analysis_result.xlsx",
        as_attachment=True
    )

if __name__ == '__main__':
    app.run(debug=True)