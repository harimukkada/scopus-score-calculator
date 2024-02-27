import os
import tkinter as tk
from pybliometrics.scopus import ScopusSearch, SerialTitle
import pandas as pd
import re

def extract_percentile(rank):
    if isinstance(rank, list):
        rank_text = str(rank[0])
    else:
        rank_text = str(rank)

    matches = re.findall(r'percentile=(\d+)', rank_text)
    return matches[0] if matches else None

def extract_and_find_position(authors_info, scopus_author_id):
    matches = re.findall(r'\d+', authors_info)
    
    if matches:
        for idx, author_id in enumerate(matches, start=1):
            if author_id == scopus_author_id:
                return str(idx)

    return 'N/A'

def get_authors_info(publication):
    authors = getattr(publication, 'authors', None) or getattr(publication, 'author_ids', None)
    if authors:
        if isinstance(authors, list):
            return "; ".join(f"Scopus Author ID - {author.auid}" for author in authors)
        else:
            return f"Scopus Author ID - {authors}"
    else:
        return ""

def get_publication_year(publication):
    if hasattr(publication, 'coverDate'):
        return publication.coverDate.split('-')[0]
    else:
        return 'N/A'

def citation_count_to_text(citation_count):
    if pd.notna(citation_count):
        return f"{citation_count:,}"  # Format as text with commas
    else:
        return 'N/A'

def calculate_citation_points(citation_count):
    if pd.notna(citation_count) and int(citation_count.replace(',', '')) > 5:
        return 10
    elif pd.notna(citation_count) and int(citation_count.replace(',', '')) <= 5:
        return 8
    else:
        return 'N/A'

def export_publications_to_excel(scopus_author_id, success_label):
    search = ScopusSearch(f"AU-ID({scopus_author_id})")
    
    if search.results is None:
        success_label.config(text="No results found for the provided author ID.")
        return
    
    data = []

    for idx, publication in enumerate(search.results, start=1):
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
            idx, authors_info, title, source_type, citation_count, doi, serial_title,
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
        'Publication', 'Authors', 'Title', 'Source Type', 'Citation Count', 'DOI', 'Serial Title',
        'Publisher', 'ISSN', 'E-ISSN', 'Scopus ID',
        'Year', 'CiteScore', 'Status', 'Document Count', 'Citation Count', 'Percent Cited', 'Rank',
        'Publication Year', 'Citation Point'  # New column
    ]

    df = pd.DataFrame(data, columns=columns)
    df['Percentile'] = df['Rank'].apply(extract_percentile)
    df['Scopus_Author_ID_Position'] = df.apply(lambda row: extract_and_find_position(row['Authors'], scopus_author_id), axis=1)

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
    df['Mark Pool'] = df['Scopus_Author_ID_Position'].apply(lambda x: 2 if x == '1' else 1)

    # Additional column for total authors based on Scopus_Author_ID_Position
    df['Total Authors'] = df['Authors'].apply(lambda x: len(re.findall(r'\d+', x)))
    
    # Additional column for co-authors (subtracting 1 from Total Authors)
    df['Co Authors'] = df['Total Authors'] - 1

    # Additional column for total points based on the provided formula
    df['Total Points'] = df.apply(lambda row: 
        row['Marks'] * row['Mark Pool'] / row['Co Authors'] if pd.notna(row['Marks']) and pd.notna(row['Mark Pool']) and row['Co Authors'] > 0 else 'N/A',
        axis=1
    )

    # Get the default Downloads folder path
    downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Specify the path of the Excel file
    excel_filename = os.path.join(downloads_folder, f"publications_info_{scopus_author_id}.xlsx")
    
    # Save the DataFrame to Excel
    df.to_excel(excel_filename, index=False)
    
    # Update success label text
    success_label.config(text=f"Excel file saved: {excel_filename}")

    # Print success message
    print("Data exported successfully!")

def on_button_click():
    scopus_author_id = entry.get()
    export_publications_to_excel(scopus_author_id, success_label)

# Create a basic Tkinter window
root = tk.Tk()
root.title("Scopus Data Exporter")

# Get the screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Set the window size and position
window_width = int(screen_width * 0.5)
window_height = int(screen_height * 0.5)
window_x = (screen_width - window_width) // 2
window_y = (screen_height - window_height) // 2

# Set the window size and position
root.geometry(f"{window_width}x{window_height}+{window_x}+{window_y}")

entry_label = tk.Label(root, text="Enter Scopus Author ID:")
entry_label.pack()

entry = tk.Entry(root)
entry.pack()

button = tk.Button(root, text="Export Data", command=on_button_click)
button.pack()

# Label to display success message
success_label = tk.Label(root, text="")
success_label.pack()

# Run the Tkinter event loop
root.mainloop()
