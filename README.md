# MSWord_ToggleImagesVisibility
A VBA macro for Microsoft Word that toggles the visibility of all images in a document. This was created as a workaround after Microsoft removed the `Show Image Placeholders` option from recent versions of Word. The goal was to provide a fast, non‑destructive way to hide or reveal images—especially useful when working with sensitive or graphic content.

## Functionality

The macro toggles visibility for two types of images:

 - Inline images: stored in `ActiveDocument.InlineShapes`, controlled by `Range.Font.Hidden`
 - Floating images: stored in `ActiveDocument.Shapes`, controlled by `Shape.Visible`

It loops through each image in the document and flips its visibility state. If all images are visible, running the macro hides them; if all are hidden, running it reveals them. Because these two types behave differently, the macro handles each separately:

 - Inline images are "hidden" by toggling the `Hidden` font property on their range.
 - Floating images are hidden by toggling the `Visible` property of the shape object.

This approach is fully reversible and does not modify or delete the images themselves.


## Limitations

 - The macro assumes that all images start in the same visibility state. If some are already hidden manually, the toggle may produce mixed results.
 - It does not process images in headers or footers. This was intentional, as the macro was designed for scenarios where the user needs to hide images in the main editing area—for example, when working with graphic medical content.

## Why This Exists

This macro fills a gap left by the removal of Word’s built‑in "Show Image Placeholders" feature. It provides a native, script‑based way to hide images quickly without altering the document’s structure or content.
