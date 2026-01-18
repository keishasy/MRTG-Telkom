import requests
import re
import calendar

class CactiEngine:
    def __init__(self):
        self.base_urls = [
            ("http://10.62.8.136/cacti", "136"),
            ("http://10.62.8.135/cacti", "135")
        ]

    def get_date_range(self, month_name, year):
        month_idx = list(calendar.month_name).index(month_name)
        last_day = calendar.monthrange(year, month_idx)[1]
        start_date = f"{year}-{month_idx:02d}-01"
        end_date = f"{year}-{month_idx:02d}-{last_day:02d}"
        return start_date, end_date

    def search_graphs(self, sid, month_name, year):
        start_d, end_d = self.get_date_range(month_name, year)
        results = []

        for base, label in self.base_urls:
            try:
                resp = requests.get(
                    f"{base}/graph_view.php",
                    params={
                        "action": "preview",
                        "filter": sid,
                        "date1": start_d,
                        "date2": end_d
                    },
                    timeout=5
                )

                graph_ids = set(re.findall(r'local_graph_id=(\d+)', resp.text))
                for gid in graph_ids:
                    results.append({
                        "url": f"{base}/graph_image.php?action=view&local_graph_id={gid}&rra_id=3",
                        "server": label
                    })
            except:
                continue

        return results
