Excel interweave-intercalating cells:

```vbscript
=INDEX($A$1:$B$5,QUOTIENT(ROW()+1,2),IF(MOD(ROW(),2)=0,2,1))
```


#Label points on a density plot with the max value of the density curve.
<p>
<img src="images/density.png">
 </p>


```r
p<-ggplot(product,aes(x=value, group=analyte, fill=analyte)) +
geom_density(adjust=1.5, alpha=.4) +
scale_x_log10()+
  theme(legend.position="none",
     axis.text=element_text(size = 12, color = "black", face = "bold.italic"))+    
ylab("density") +
xlab("Analyte Value (log10)") %>% 

###pull out the numbers from the generated graph to be able to plot the values.

 ggplot_build() %>% 
  `[[`("data") %>% 
  `[[`(1) %>% 
  as_tibble() %>% 
  group_by(group) %>% 
summarise(max=max(y))r


```
