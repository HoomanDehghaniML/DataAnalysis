import pandas as pd

base = pd.read_csv(
    r"C:\Users\Work\Downloads\Competitor rankings.csv",
    index_col="Company"
)

# Make a list of competitors that rank 1
rank_1s = []

for label, column in base.items():
    for index, value in column.items():
        if value == 1:
            rank_1s.append({label: index})


# Find overall strongest competitors
median_ascending = (base.median(axis=1)
                        .sort_values(ascending=True)
                        .to_frame()
                        .rename(columns={0: "Median"}))
median_ascending_html = median_ascending.to_html()

# Make the final list
# Final list includes the top 8 of the list and the 1st list minus duplicates
top_6 = []
counter = 1
for index, value in base.median(axis=1).items():
    if counter <= 6:
        top_8.append(index)
    else:
        break
    counter += 1

# Write to html file and paste it next to the QC Report
with open(
    r"G:\Shared drives\sq1 - marketing\seo\Reports\
      Product Page Analysis\QC q4 2022\Competitors.html", 'w'
) as file:
    file.write(
        f"""<p>
        <strong>Competitors that have ranked 1st for a keyword:</strong><br>
        {rank_1s[0:3]}<br>
        {rank_1s[3:]}
        </p>"""
    )
    file.write(
        median_ascending_html
    )
