from . import Cell_text_Formatting , Cell_text_insertion, Dataframe_mapping , Table_insertion_and_formatting , Risk_map , aryan
from .aryan import (update_page_no, fill_detailed_observation_templates)
from .Cell_text_Formatting import _parse_color , format_text_in_table_cell
from .Table_insertion_and_formatting import replace_rectangle_with_table_and_handle_overflow , ann_replace_rectangle_with_table_and_handle_overflow
from .Cell_text_insertion import (duplicate_slide , split_text_by_max_lines , handle_cell_overflow , format_text_in_shape , check_overflow_text , fill_template_with_header_map_and_title,  fill_templates_with_header_map_and_title_v2)
from .Dataframe_mapping import make_ordered_hashmap , map_columns , _find_table_for_mapping , remove_whitespace_cells

__all__ = ["Cell_text_Formatting" , "Cell_text_insertion" , "Dataframe_mapping" ,
            "_parse_color" , "format_text_in_table_cell" , "duplicate_slide" , "split_text_by_max_lines" , 
            "handle_cell_overflow" , "format_text_in_shape" , "check_overflow_text" , "fill_template_with_header_map_and_title" , 
            "make_ordered_hashmap" , "map_columns" , "_find_table_for_mapping" , "remove_whitespace_cells" , "replace_rectangle_with_table_and_handle_overflow"  , "ann_replace_rectangle_with_table_and_handle_overflow",
            "update_page_no",
            "fill_templates_with_header_map_and_title_v2",
            "fill_detailed_observation_templates"]