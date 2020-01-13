from Data_Manager import ConditionalFormat


def cell_shading(
        bg_color: str,
        font_color: str,
        criteria: str,
        value: str = None,
        min: str = None,
        max: str = None
) -> ConditionalFormat:
    if value:
        return ConditionalFormat(
            {
                'bg_color': bg_color,
                'dont_color': font_color
            },
            criteria,
            value=value
        )
    else:
        return ConditionalFormat(
            {
                'bg_color': bg_color,
                'dont_color': font_color
            },
            criteria,
            minimum=min,
            maximum=max
        )
