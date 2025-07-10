def adjust_font_size(
    text: str,
    max_width_px: int,
    ax,
    default_size: int = 12,
    min_size: int = 8
) -> int:
    """
    动态调整字体大小，使 text 在 ax 上渲染时宽度不超过 max_width_px
    """
    if len(str(text)) <= 22:
        return default_size

    fig = ax.figure
    fig.canvas.draw()  # 确保 renderer 准备好
    renderer = fig.canvas.get_renderer()

    for size in range(default_size, min_size - 1, -1):
        text_obj = ax.text(0, 0, text, fontsize=size, weight="bold")
        bbox = text_obj.get_window_extent(renderer=renderer)
        text_obj.remove()
        if bbox.width <= max_width_px:
            return size

    return min_size
