from flask import Flask, render_template, request, send_file
from io import BytesIO
import pandas as pd
from pybliometrics.scopus import ScopusSearch, SerialTitle

app = Flask(__name__)

def export_publications_to_excel(scopus_author_id):
    search = ScopusSearch(f"AU-ID({scopus_author_id})")
    if search.results is None:
        return None

    data = []

    for publication in search.results:
        authors_info = get_authors_info(publication)
        title = publication.title
        source_type = publication.subtypeDescription
        citation_count = citation_count_to_text(publication.citedby_count)  # Format citation count as text
        doi = publication.doi

        try:
            serial_id = publication.issn or publication.eIssn
            source_full = SerialTitle(serial_id, view="CITESCORE")
            serial_title = source_full.title
        except Exception as e:
            if 'The resource specified cannot be found' in str(e):
                source_full = None
                serial_title = "N/A"
            else:
                raise

        row = [
            authors_info, title, source_type, citation_count, doi, serial_title,
            getattr(source_full, 'publisher', 'N/A'),
            getattr(source_full, 'issn', 'N/A'),
            getattr(source_full, 'eissn', 'N/A'),
            getattr(source_full, 'source_id', 'N/A'),
            (source_full.citescoreyearinfolist[0].year if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A',
            (source_full.citescoreyearinfolist[0].citescore if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A',
            (source_full.citescoreyearinfolist[0].status if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A',
            (source_full.citescoreyearinfolist[0].documentcount if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A',
            (source_full.citescoreyearinfolist[0].citationcount if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A',
            (source_full.citescoreyearinfolist[0].percentcited if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A',
            (source_full.citescoreyearinfolist[0].rank if source_full.citescoreyearinfolist else 'N/A') if source_full else 'N/A'
        ]
        
        # Additional column for publication year
        row.append(get_publication_year(publication))
        
        # Additional column for citation points
        row.append(calculate_citation_points(citation_count))

        data.append(row)

    columns = [
        'Authors', 'Title', 'Source Type', 'Citation Count', 'DOI', 'Serial Title',
        'Publisher', 'ISSN', 'E-ISSN', 'Scopus ID',
        'Year', 'CiteScore', 'Status', 'Document Count', 'Citation Count', 'Percent Cited', 'Rank',
        'Publication Year', 'Citation Point'  # New column
    ]

    df = pd.DataFrame(data, columns=columns)
    df['Percentile'] = df['Rank'].apply(extract_percentile)

    # Additional column for marks with updated conditions
    df['Marks'] = df.apply(lambda row: 
        50 if row['Percentile'] is not None and int(row['Percentile']) >= 90 else
        25 if row['Percentile'] is not None and 75 <= int(row['Percentile']) <= 89 else
        12.5 if row['Percentile'] is not None and 50 <= int(row['Percentile']) <= 74 else
        10 if row['Percentile'] is not None and int(row['Percentile']) < 50 else
        pd.notna(row['Citation Point']) and row['Citation Point'] or 'N/A',
        axis=1
    )

    # Fill N/A values in the "Marks" column with values from "Citation Point" column
    df['Marks'].fillna(df['Citation Point'], inplace=True)

    # Additional column for quartile based on marks
    df['Quartile'] = df['Marks'].apply(lambda x: 
        'Q1' if x == 50 else
        'Q2' if x == 25 else
        'Q3' if x == 12.5 else
        'Q4' if x == 10 else ''
    )

    # Additional column for mark pool based on Scopus_Author_ID_Position
    df['Mark Pool'] = 1

    # Additional column for total authors based on Scopus_Author_ID_Position
    df['Total Authors'] = df['Authors'].apply(lambda x: len(re.findall(r'\d+', x)))
    
    # Additional column for co-authors (subtracting 1 from Total Authors)
    df['Co Authors'] = df['Total Authors'] - 1

    # Additional column for total points based on the provided formula
    df['Total Points'] = df.apply(lambda row: 
        row['Marks'] * row['Mark Pool'] / row['Co Authors'] if pd.notna(row['Marks']) and pd.notna(row['Mark Pool']) and row['Co Authors'] > 0 else 'N/A',
        axis=1
    )

    # Save the DataFrame to Excel
    excel_buffer = BytesIO()
    df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)

    return excel_buffer

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        scopus_author_id = request.form["scopus_author_id"]
        excel_buffer = export_publications_to_excel(scopus_author_id)
        if excel_buffer:
            return send_file(
                excel_buffer,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                attachment_filename=f"publications_info_{scopus_author_id}.xlsx"
            )
        else:
            return "No results found for the provided author ID."
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
