#' Convert to UNHCR branded Powerpoint presentation
#'
#' Format for converting from R Markdown to an UNHCR branded Powerpoint presentation
#'
#' @rdname pptx_presentation
#'
#' @importFrom officedown rpptx_document
#'
#' @param ... extra parameters to pass to `rmarkdown::powerpoint_presentation`
#'
#' @export
pptx_presentation <- function(...) {

  ref_doc <- pkg_resource("pptx/pptx_reference.pptx")

  officedown::rpptx_document(
    reference_doc = ref_doc,
    master = "UNHCR_PPTX",
    ...
  )
}

#' @rdname pptx_presentation
#' @noRd
ppt_presentation <- pptx_presentation
