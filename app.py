from flask import Flask, render_template, request, send_file
import pandas as pd
import io
import os

app = Flask(__name__)

last_matches = []

def get_score(rank_index, choices_len):
    return choices_len - rank_index


# ================== MATCHING ==================

def mutual_matching(men_choices, women_choices):
    pairs = []

    for m in range(len(men_choices)):
        for i, w in enumerate(men_choices[m]):
            if 0 <= w < len(women_choices) and m in women_choices[w]:
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


def fallback_matching(men_choices, women_choices, matched_men, matched_women):
    pairs = []

    for m in range(len(men_choices)):
        if m in matched_men:
            continue

        for i, w in enumerate(men_choices[m]):
            if w in matched_women or not (0 <= w < len(women_choices)):
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


# ================== ROUTES ==================

@app.route("/", methods=["GET", "POST"])
def index():
    global last_matches

    matches = []
    men_choices = []
    women_choices = []
    n = 0
    error = ""

    if request.method == "POST":

        num_people = request.form.get("num_people")
        n_value = request.form.get("n")

        # STEP 1
        if num_people:
            try:
                n = int(num_people)
            except ValueError:
                n = 0
            return render_template("index.html", n=n)

        # STEP 2
        elif n_value:
            try:
                n = int(n_value)
            except ValueError:
                n = 0

            for i in range(n):
                men_data = request.form.get(f"men_{i}")
                women_data = request.form.get(f"women_{i}")

                if not men_data or not women_data:
                    error = "Semua peserta harus diisi"
                    break

                try:
                    men = [int(x) - 1 for x in men_data.split()]
                    women = [int(x) - 1 for x in women_data.split()]
                except ValueError:
                    error = "Input harus berupa angka"
                    break

                # VALIDASI
                if any(x < 0 or x >= n for x in men) or any(x < 0 or x >= n for x in women):
                    error = f"ID harus antara 1 sampai {n}"
                    break

                if len(set(men)) != len(men) or len(set(women)) != len(women):
                    error = "Tidak boleh ada duplikat"
                    break

                men_choices.append(men)
                women_choices.append(women)

            if len(men_choices) != n:
                return render_template("index.html", n=n, error=error)

            # MATCHING
            m1, matched_men, matched_women = mutual_matching(men_choices, women_choices)
            m2 = fallback_matching(men_choices, women_choices, matched_men, matched_women)

            # Cek pria dan wanita secara terpisah dengan if + if (bukan if/elif)
            unmatched = []
            for idx in range(n):
                if idx not in matched_men:
                    unmatched.append((idx + 1, "-", "-", "Tidak Terpilih"))
                if idx not in matched_women:
                    unmatched.append(("-", idx + 1, "-", "Tidak Terpilih"))

            matches = [(m+1, w+1, score, tipe) for (m, w, score, tipe) in (m1 + m2)]
            matches += unmatched
            last_matches = matches

            # Konversi choices ke 1-based untuk ditampilkan
            men_choices_display = [[x + 1 for x in row] for row in men_choices]
            women_choices_display = [[x + 1 for x in row] for row in women_choices]

            return render_template(
                "index.html",
                matches=matches,
                men_choices=men_choices_display,
                women_choices=women_choices_display,
                n=n,
                error=error
            )

    return render_template(
        "index.html",
        matches=matches,
        men_choices=men_choices,
        women_choices=women_choices,
        n=n,
        error=error
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


# ================== RUN ==================

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)