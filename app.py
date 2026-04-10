from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import os

app = Flask(__name__)

last_matches = []

def get_score(rank_index, choices_len):
    return choices_len - rank_index

def mutual_matching(men_choices, women_choices, n):
    pairs = []

    for m in range(n):
        for i, w in enumerate(men_choices[m]):
            if m in women_choices[w]:
                j = women_choices[w].index(m)
                score = get_score(i, len(men_choices[m])) + get_score(j, len(women_choices[w]))
                pairs.append((score, m, w))

    pairs.sort(reverse=True)

    matched_men = set()
    matched_women = set()
    matches = []

    for score, m, w in pairs:
        if m not in matched_men and w not in matched_women:
            matched_men.add(m)
            matched_women.add(w)
            matches.append((m, w, score, "Mutual"))

    return matches, matched_men, matched_women

def fallback_matching(men_choices, women_choices, matched_men, matched_women, n):
    pairs = []

    for m in range(n):
        if m in matched_men:
            continue
        for i, w in enumerate(men_choices[m]):
            if w in matched_women:
                continue

            if m in women_choices[w]:
                j = women_choices[w].index(m)
                score = get_score(i, len(men_choices[m])) + get_score(j, len(women_choices[w]))
            else:
                score = get_score(i, len(men_choices[m]))

            pairs.append((score, m, w))

    pairs.sort(reverse=True)

    matches = []

    for score, m, w in pairs:
        if m not in matched_men and w not in matched_women:
            matched_men.add(m)
            matched_women.add(w)
            matches.append((m, w, score, "Fallback"))

    return matches

@app.route("/", methods=["GET", "POST"])
def index():
    global last_matches

    matches = []
    men_choices = []
    women_choices = []
    n = 0

    if request.method == "POST":

        # Ambil dua kemungkinan input
        num_people = request.form.get("num_people")
        n_value = request.form.get("n")

        # STEP 1: input jumlah peserta
        if num_people:
            try:
                n = int(num_people)
            except:
                n = 0
            return render_template("index.html", n=n)

        # STEP 2: proses matching
        elif n_value:
            try:
                n = int(n_value)
            except:
                n = 0

            for i in range(n):
                men_data = request.form.get(f"men_{i}")
                women_data = request.form.get(f"women_{i}")

                # VALIDASI biar tidak None
                if not men_data or not women_data:
                    continue

                men = list(map(int, men_data.split()))
                women = list(map(int, women_data.split()))

                men_choices.append(men)
                women_choices.append(women)

            m1, matched_men, matched_women = mutual_matching(men_choices, women_choices, n)
            m2 = fallback_matching(men_choices, women_choices, matched_men, matched_women, n)

            matches = m1 + m2
            last_matches = matches

    return render_template(
        "index.html",
        matches=matches,
        men_choices=men_choices,
        women_choices=women_choices,
        n=n
    )

@app.route("/export")
def export():
    global last_matches

    df = pd.DataFrame(last_matches, columns=["Pria", "Wanita", "Skor", "Tipe"])

    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(output,
                     download_name="matching_result.xlsx",
                     as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    app.run(host="0.0.0.0", port=port)