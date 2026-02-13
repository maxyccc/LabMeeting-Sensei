from src.slide_assembly import create_presentation, insert_slide, save_presentation
from src.io import search_pattern_file
from src.basic_layouts import HorizontalLayout, SingleImageLayout, SingleTextLayout, VerticalLayout, AdaptiveGridLayout



def get_example_animal_figs(index: int):
    all_matches = search_pattern_file(pattern=f"{index}.jpg", search_dir=".")
    assert len(all_matches) == 1, f"Expected 1 match for {index}, got {len(all_matches)}"
    return SingleImageLayout(all_matches[0], padding=2)

if __name__ == "__main__":
    top_row = HorizontalLayout([
        get_example_animal_figs(1),
        get_example_animal_figs(2),
        get_example_animal_figs(3),
    ])
    center_top = HorizontalLayout([
        get_example_animal_figs(6),
        get_example_animal_figs(7),
        VerticalLayout([
            get_example_animal_figs(8),
            get_example_animal_figs(9),
        ])
    ])
    center_bottom = HorizontalLayout([
        get_example_animal_figs(10),
        VerticalLayout([
            get_example_animal_figs(11),
            get_example_animal_figs(12),
        ]),
        get_example_animal_figs(13),
    ])
    center = HorizontalLayout([
        get_example_animal_figs(4),
        VerticalLayout([
            center_top,
            center_bottom,
        ]),
        get_example_animal_figs(5),    
    ])
    final_layout = VerticalLayout([
        top_row,
        center,
        get_example_animal_figs(14),
    ])
    final_layout.render().save("example_result.png")

    ppt = create_presentation()
    insert_slide("example_result.png", ppt)
    save_presentation(ppt, "./example_result.pptx")
