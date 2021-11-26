#' Convert to UNHCR branded Word document
#'
#' Format for converting from R Markdown to an UNHCR branded Word document
#'
#' @rdname docx_document
#'
#' @importFrom officedown rdocx_document
#'
#' @param ... extra parameters to pass to `officedown::rdocx_document`
#'
#' @export
docx_document <- function(...) {

  ref_doc <- pkg_resource("docx/docx_reference.docx")

  tables <- list(
    style = "Plain Table 31",
    layout = "autofit",
    width = 1,
    caption = list(style = "Table Caption", pre = "Table ", sep = ": "),
    conditional = list(
      first_row = TRUE, first_column = FALSE, last_row = FALSE,
      last_column = FALSE, no_hband = FALSE, no_vband = TRUE
    )
  )

  plots <- list(
    style = "Normal",
    align = "center",
    caption = list(
      style = "Image Caption",
      pre = "Figure ",
      sep = ": "
    )
  )

  officedown::rdocx_document(
    reference_docx = ref_doc,
    tables = tables,
    plots = plots,
    ...
  )

}