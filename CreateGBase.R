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
  rename(Generation =Generation.B.....Year.Round...Transport.Model.) %>%#rename generation for ease of taking away negative demand
  mutate(Generation = if_else(Demand<0, Generation-Demand, Generation),
         Demand = if_else(Demand<0, 0, Demand)) %>%
  group_by(Name) %>%
  summarise(Voltage = max(Voltage),
            Demand = sum(Demand),
            Generation = sum(Generation),
            #Peak_Gen = sum(Generation.A.....Peak.Security...Transport.Model.),
            BalencedPower = sum(BusTransferB),
            Bus.Order = min(Bus.Order)) 

trans2<- TransportData[,17:59] %>%
  setNames(make.names(names(.))) %>% as_tibble %>%
  filter(!is.na(Bus.1)) %>% #remove emty rows
  mutate(Bus.1 = str_sub(Bus.1, 1, 4) %>% gsub("-|_", "",.),   #convert to site from individual nodes, this simplifies a lot a prevents sites being taken down piecemeal
         Bus.2 = str_sub(Bus.2, 1, 4)%>% gsub("-|_", "",.)) %>%
  #converting into an undirected graph and back afain puts all the node names in alphanumeric order, this prevents
  #The grouping from not correctly aggregating due to node pairs such as B to A followed by from A to B. This operation makes all
  #from A to B.
  graph_from_data_frame(., directed = FALSE) %>%
  as_data_frame()  %>%
  rename(Bus.1 = from, Bus.2 = to) %>%
  filter(!(Bus.1==Bus.2))%>% #removes internal connections, about 400 edges
  group_by(Bus.1, Bus.2) %>% #trying to stop the non-unique identifier problem
  summarise(
    Y = sum(1/X..Peak.Security.), #create susceptance, The addmitance of parallel lines is Y1+Y2+Y3+..+Yn
    Link.Limit = sum(Link.Limit),
    Length = mean(OHL.Length + Cable.Length), #They are all pretty much the same length when they match like this, so this isn't very important
    Combines = n() #tells howmany parallel edges between two different sites were combined together, is about 300 edges
  ) %>% 
  ungroup %>%
  mutate(Link = paste(Bus.1,Bus.2, sep = "-"),
         #These f lines have a lower limit than thier initial values. They are given the median value for thier voltage
         Link.Limit = case_when(
           Link == "BONB-BRAC" ~1090,
           Link == "FAUG-LAGG" ~329, #This is just an arbitrary number as this line is loaded so much more than median
           Link == "KEIT-KINT" ~1090,
           Link == "LAGG-MILW" ~203,
           TRUE ~ Link.Limit
         )
  )

#make sure everything is under limit
# as_data_frame(gbase) %>%
#   group_by(Voltage) %>%
#   summarise(mean = mean(Link.Limit),
#             median = median(Link.Limit))
# 
# as_data_frame(gbase) %>%
#   filter(Link.Limit<PowerFlow)

gbase <- graph_from_data_frame(trans2, directed=FALSE, vertices = VertexMetaData)

gbase <- set.edge.attribute(gbase, "name", value = get.edge.attribute(gbase, "Link")) %>%
  set.edge.attribute(., "weight", value = get.edge.attribute(gbase, "Link.Limit")) %>% #weight should probably be the admittance but whatever.
  set.vertex.attribute(., "component", value = components(gbase)$membership)


SlackRef <- SlackRefFunc(gbase) #find the most appropriate node to be the slack bus

#Add in Edge Voltages
gbase <- set_edge_attr(gbase, "Voltage", 
                       #Takes the minimum Voltage when the nodes do not agree
                       #If statement only necessary if you need to take a value that is neither max nor min, E.g. 0
                       value = pmin(get.vertex.attribute(gbase, "Voltage", get.edgelist(gbase)[,1]), 
                                    get.vertex.attribute(gbase, "Voltage", get.edgelist(gbase)[,2]))
                       ) %>%
  PowerFlow(., SlackRef$name) #calculate power flow
