setwd(Tariff)

GenerationData <- read_excel("Tariff & Transport Model_2018_19 Tariffs_External.xlsm", sheet =10, skip = 33) %>%
  setNames(make.names(names(.))) %>%
  mutate(Site = str_sub(Node.1,1,4))

LocalAssetData <- read_excel("Tariff & Transport Model_2018_19 Tariffs_External.xlsm", sheet =11, skip = 11) %>%
  setNames(make.names(names(.)))

TransportData <- read_excel("Tariff & Transport Model_2018_19 Tariffs_External.xlsm", sheet =12, skip = 11) %>%
  setNames(make.names(names(.)))

#Clean and organise the data
trans1 <-TransportData[,1:16] %>% 
  filter(!is.na(Bus.ID)) %>%
  mutate(Bus.Name = gsub("-", "_", Bus.Name)) #haveing "-" means that edge names become inseperable. thus all "-" are converted to "_"

VertexMetaData <- trans1 %>%
  mutate(Name = str_sub(Bus.Name, 1, 4)%>% gsub("-|_", "",.)) %>% #converts to site level for nodes, this reduces by about 450 nodes almost half
  group_by(Name) %>%
  summarise(Voltage = max(Voltage),
            Demand = sum(Demand),
            Generation = sum(Generation.B.....Year.Round...Transport.Model.),
            BalencedPower = sum(BusTransferB),
            Bus.Order = min(Bus.Order)) 


trans2 <- TransportData[,17:59] %>%
  setNames(make.names(names(.))) %>% 
  filter(!is.na(Bus.1)) %>% #remove emty rows
  mutate(Bus.1 = str_sub(Bus.1, 1, 4) %>% gsub("-|_", "",.),   #convert to site from individual nodes, this simplifies a lot a prevents sites being taken down piecemeal
         Bus.2 = str_sub(Bus.2, 1, 4)%>% gsub("-|_", "",.)) %>%
  filter(!(Bus.1==Bus.2)) %>% #removes internal connections, about 400 edges
  group_by(Bus.1, Bus.2) %>% #trying to stop the non-unique identifier problem
  summarise(
    Y = mean(1/X..Peak.Security.), #create susceptance
    Link.Limit = sum(Link.Limit),
    Length = mean(OHL.Length + Cable.Length), #They are all pretty much the same length when they match like this, so this isn't very important
    Combines = n() #tells howmany parallel edges between two different sites were combined together, is about 300 edges
  ) %>% 
  ungroup %>%
  mutate(Link = paste(Bus.1,Bus.2, sep = "-"))


gbase <- graph_from_data_frame(trans2, directed=FALSE, vertices = VertexMetaData)

gbase <- set.edge.attribute(gbase, "name", value = get.edge.attribute(gbase, "Link")) %>%
  set.edge.attribute(., "weight", value = get.edge.attribute(gbase, "Link.Limit")) %>%
  set.vertex.attribute(., "component", value = components(gbase)$membership)


#Add in Edge Voltages
gbase <- set_edge_attr(gbase, "Voltage", 
                       #Takes the minimum Voltage when the nodes do not agree
                       #If statement only necessary if you need to take a value that is neither max nor min, E.g. 0
                       value = pmin(get.vertex.attribute(gbase, "Voltage", get.edgelist(gbase)[,1]), 
                                    get.vertex.attribute(gbase, "Voltage", get.edgelist(gbase)[,2]))
                       ) %>%
  PowerFlow(., SlackRef = "GLLE" )

#Finds edges with no powerflow
UselessEdge <- (get.edge.attribute(gbase, "PowerFlow")==0)
print(get.edge.attribute(gbase, "name", index = (1:length(UselessEdge))[UselessEdge]) )
#Remove Edges that are not doing anything and then any islanded nodes.
#This is done before the simulation begins
gbase <- delete.edges(gbase, (1:length(UselessEdge))[UselessEdge]) %>%
  BalencedGenDem(., "Demand", "Generation")
rm(UselessEdge)


