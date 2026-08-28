pl <- ggplot(data.frame(x = 1), aes(x,x)) + geom_text(aes(label = "File does not exist"), fontface = "bold") + 
  theme_void() + 
  theme(panel.background = element_rect(fill = "lightblue")) + 
  theme(text = element_blank()) + 
  geom_blank()
pl

flplot <- "C:/PROJECTS/1_PROJTOOLS/plotExcel/inst/fileDoesNotExist.pdf"
ggsave(plot = pl, filename = flplot, width = 5, height = 3, units = "cm")

system(paste('"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"', shQuote(normalizePath(flplot))), wait = FALSE)
