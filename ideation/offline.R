filedata <- officer::read_docx("D:\\light.docx")

summarized <- officer::docx_summary(filedata)

toc <- summarized %>%
  dplyr::filter(content_type == "paragraph" & tolower(style_name) %in% c('heading 1','heading 2')) %>%
  mutate(index = 1:n()) %>%
  mutate(subtext = ifelse(tolower(style_name) == "heading 1", text, NA)) %>%
  tidyr::fill(subtext) %>%
  dplyr::select(text, index, style_name, subtext)

sumdoc_indexed <- summarized %>%
  left_join(toc) %>%
  tidyr::fill(index)


input_sections <- c("Copanlisib (BAY 80-6946)","Synopsis")


### Download



temp <-sumdoc_indexed %>%
  #filter(text %in% input_sections) %>%
  slice(match(input_sections, text)) %>%
  mutate(key = 1:n()) %>%
  select(index, key) %>%
  inner_join(sumdoc_indexed)

temp2 <- temp %>% 
  mutate(text = ifelse(style_name == "List Paragraph", paste("*",text), text)) %>%
  filter(!is.na(text) | text != "")  

out <- paste(temp2$text, collapse = "\n\n")

text <- out %>% strsplit('\n\n') %>% unlist()

doc <- officer::read_docx("template.docx")