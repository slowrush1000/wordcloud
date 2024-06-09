
import sys
import pandas as pd
from googlesearch import search

# Search query and date range
query = "Raisi"
start_date = "2024-05-19"
end_date = "2024-05-28"

# Perform Google News search
search_results = search(f"{query} site:news.google.com", lang='en', advanced=True)

# Extract headlines
#headlines = [result for result in search_results]
#print(search_results)
for result in search_results:
    print([result for result in search_results])
#    print(f'{result.title} {result.description}')

# Create a DataFrame
#df = pd.DataFrame({"Titles:": headlines})

# Save to Excel
#output_file = "Raisi_headlines.xlsx"
#df.to_excel(output_file, index=False)

#print(f"Headlines saved to {output_file}")