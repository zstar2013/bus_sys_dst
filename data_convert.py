remove_route=lambda x:x.split("路")[-1].strip()
remove_line=lambda x:x.split("线")[-1].strip()
remove_diagonal=lambda x:x.split("/")[-1].strip()