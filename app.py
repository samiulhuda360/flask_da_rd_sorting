from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
from openpyxl import Workbook
from urllib.parse import urlparse
from openpyxl.utils.cell import get_column_letter

import csv
import io

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Set a secret key for Flash messages

# Define the rating ranges
rating_ranges = [
    (0, 10),
    (11, 20),
    (21, 30),
    (31, 40),
    (41, 50),
    (51, 60),
    (61, 70),
    (71, 80),
    (81, 90),
    (91, 100)
]

# Define the referring domains ranges
rd_ranges = [
    (1, 100),
    (101, 200),
    (201, 300),
    (301, 400),
    (401, 500),
    (501, 600),
    (601, 700),
    (701, 800),
    (801, 900),
    (901, 1000),
    (1001, float('inf'))  # 1000+
]

@app.route('/domain-sorting', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('files[]')

        if not uploaded_files:
            flash('No file chosen. Please select at least one file.', 'error')
            return redirect(url_for('index'))

        if len(uploaded_files) > 5:
            flash('Too many files uploaded. Maximum allowed is 5.', 'error')
            return redirect(url_for('index'))

        rating_counts = {}

        for file in uploaded_files:
            if file.filename.endswith('.csv'):
                file_content = file.stream.read().decode('utf-8')
                csv_reader = csv.reader(file_content.splitlines())

                # Read the header row
                header_row = next(csv_reader)

                # Find the index of the "Domain rating", "Target URL", and "Referring domains" columns
                domain_rating_column_index = None
                target_url_column_index = None
                referring_domains_column_index = None
                for index, column_name in enumerate(header_row):
                    if "Domain rating" in column_name:
                        domain_rating_column_index = index
                    elif "Target URL" in column_name:
                        target_url_column_index = index
                    elif "Referring domains" in column_name:
                        referring_domains_column_index = index

                if domain_rating_column_index is None or target_url_column_index is None or referring_domains_column_index is None:
                    flash(f"Required columns not found in {file.filename}.", 'error')
                    return redirect(url_for('index'))

                # Iterate over each row in the CSV
                for row_index, row in enumerate(csv_reader):
                    try:
                        domain_rating = float(row[domain_rating_column_index])
                        target_url = row[target_url_column_index]
                        referring_domains = int(row[referring_domains_column_index])
                        root_domain = urlparse(target_url).netloc

                        # Initialize the count for the root domain if it doesn't exist
                        if root_domain not in rating_counts:
                            rating_counts[root_domain] = {
                                'domain_rating': {range_: 0 for range_ in rating_ranges},
                                'referring_domains': {range_: 0 for range_ in rd_ranges}
                            }

                        # Increment the count for the corresponding rating range
                        for range_ in rating_ranges:
                            if range_[0] <= domain_rating <= range_[1]:
                                rating_counts[root_domain]['domain_rating'][range_] += 1
                                break

                        # Increment the count for the corresponding referring domains range
                        for range_ in rd_ranges:
                            if range_[0] <= referring_domains <= range_[1]:
                                rating_counts[root_domain]['referring_domains'][range_] += 1
                                break
                    except (IndexError, ValueError) as e:
                        flash(f"Error processing row {row_index + 1} in {file.filename}: {e}", 'error')
                        return redirect(url_for('index'))

        if not rating_counts:
            flash('No data to generate report.', 'error')
            return redirect(url_for('index'))

        # Create a new workbook and select the active sheet
        workbook = Workbook()
        sheet = workbook.active

        # Write the header row for domain rating counts
        sheet.cell(row=1, column=1, value="Target Competitor")
        for col, range_ in enumerate(rating_ranges, start=2):
            sheet.cell(row=1, column=col, value=f"DR {range_[0]}-{range_[1]}")

        # Write the data rows for domain rating counts
        row_idx = 2
        for root_domain in rating_counts.keys():
            sheet.cell(row=row_idx, column=1, value=root_domain)
            for col, range_ in enumerate(rating_ranges, start=2):
                count = rating_counts[root_domain]['domain_rating'][range_]
                sheet.cell(row=row_idx, column=col, value=count)
            row_idx += 1

        # Write the header row for referring domains counts
        start_row = row_idx + 1
        sheet.cell(row=start_row, column=1, value="Target Competitor")
        for col, range_ in enumerate(rd_ranges, start=2):
            sheet.cell(row=start_row, column=col, value=f"RD {range_[0]}-{range_[1]}")

        # Write the data rows for referring domains counts
        row_idx = start_row + 1
        for root_domain in rating_counts.keys():
            sheet.cell(row=row_idx, column=1, value=root_domain)
            for col, range_ in enumerate(rd_ranges, start=2):
                count = rating_counts[root_domain]['referring_domains'][range_]
                sheet.cell(row=row_idx, column=col, value=count)
            row_idx += 1

        # Adjust the column width
        for col_num, col_letter in enumerate(get_column_letter(2), start=2):
            sheet.column_dimensions[col_letter].width = 15

        # Save the workbook to a BytesIO object
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name='domain_rating_report.xlsx')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)