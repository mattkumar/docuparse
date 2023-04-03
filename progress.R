# goal: take content from the protocol, transport into the csr
# the protocol is a separate doc file, the csr is a separate doc file

# libs
library(dplyr)
library(tidyr)
library(officer)

# read in protocol file
filedata <- officer::read_docx("trial-protocol.docx")

# summarize the protocol file
summarized <- officer::docx_summary(filedata)

# create TOC for protocol
# this is what the user can potentially chose from on the UI
# might need to modify line 15 to other content_types based on final transcelerate template format
toc <- summarized %>%
  dplyr::filter(content_type == "paragraph" & tolower(style_name) %in% c('heading 1','heading 2')) %>%
  mutate(index = 1:n()) %>%
  mutate(subtext = ifelse(tolower(style_name) == "heading 1", text, NA)) %>%
  tidyr::fill(subtext) %>%
  dplyr::select(text, index, style_name, subtext)

# index the summarized protocol file so there's a connection between toc and summarized
sumdoc_indexed <- summarized %>%
  left_join(toc) %>%
  tidyr::fill(index)

# simulate a csr
# we'll use the same file, no issue
csr <- officer::read_docx("trial-protocol.docx")

# simulate user selecting the following sections of the protocol for inclusion in the csr
# note the ORDER of selection: informed consent first, randomisation second
input_sections <- c("Informed Consent", "Randomisation")


# using the input sections, get the corresponding text for those choices
temp <-sumdoc_indexed %>%
  slice(match(input_sections, text)) %>%
  mutate(key = 1:n()) %>%
  select(index, key) %>%
  inner_join(sumdoc_indexed) %>%
  # The next line needs to be made robust
  arrange(desc(index), desc(doc_index))

# process corresponding text
out <- paste(temp$text, collapse = "\n\n\n")

text <- out %>% 
  strsplit('\n\n') %>% 
  unlist()

# this is a keyword predefined in the csr template "Bkm1" - look at the last page 
# it "functions" like a bookmark, but isn't a true bookmark
# it's just a point to direct where to start inserting text
# we will potentially have many of these in the CSR
csr <- cursor_reach(csr, keyword = "Bkm1")

for (t in 1:length(text)) {
  body_add_par(csr, text[[t]])
}

# check results
# material printed in the right spot, in the right order
print(csr, target = "test.docx")


