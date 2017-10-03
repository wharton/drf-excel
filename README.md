# Wagtail Gridder

Wagtail Gridder is a Bootstrap 4 enabled layout for the Wagtail CMS. Grid Items are created within categories, and displayed on a Grid Index Page. The JavaScript libraries Gridder and MixItUp are included.

# Requirements

* Django >= 1.9
* Wagtail >= 1.9 (for ParentalManyToMany)
* Bootstrap >= 3 (optimized for Bootstrap 4; templates can be modified)

# Installation

    pip install drf_renderer_xls

# Settings

    WAGTAILGRIDDER_CLEAR_CACHE = False

The default Wagtail Gridder template caches the grid display area to reduce the number of queries performed. Setting `WAGTAILGRIDDER_CLEAR_CACHE = True` in your Django settings will clear the **entire** Django cache after a page is edited. This approach is necessary, as Django does not currently support deletion from the cache by pattern. Setting this to `True` will clear your cache every time you save a Wagtail page. If anyone knows of a better solution that works for Django's supported cache systems, please let us know!

# Screenshots

## Grid Index Page:

![Grid Index Page](img/grid_index_page.jpg)

## Grid Index Page, with Grid Item expanded:

![Grid Index Page, with Grid Item expanded](img/grid_index_page_expanded.jpg)

## Optional featured hero region:

![Optional featured hero region](img/featured_hero.jpg)

## Grid Item landing page:

![Grid Item landing page](img/grid_item.jpg)

## Editing a Grid Item:

![Editing a Grid Item](img/edit_grid_item.jpg)

## Editing a Grid Index Page:

![Editing a Grid Index Page](img/edit_grid_index_page.jpg)

# Contributors

* [Timothy Allen](https://github.com/FlipperPA)
* [Charles Rejonis](https://github.com/rejonis)
