import os
import part_list_loader
d = os.path.dirname(os.path.abspath(__file__))
idx = part_list_loader.load_part_list_index(d)
print("locs", len(idx))
