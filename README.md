# Running Dinner Calculator V2

This is the new and updated version of the [Running Dinner Calculator](https://courses.gistools.geog.uni-heidelberg.de/pk256/running-dinner-calculator) created by Jakob Tinapp and David Benedict for the FOSSGIS Seminar of Wintersemester 2022/23 at the geographic institute of the University of Heidelberg.

Improvements compared to the old version include:

1. Automatic delegation of double cooks in case the number of Teams is not divisible by three,
2. Full integration of the entire process into python to reduce the chance for confusion and human error,
3. Automated Mail writing,
4. Reduced requirements (two excel spreadsheets are needed instead of one and a .kml Layer in Case of the old version).


## Description

The organizational process of a Running Dinner can get complicated especially if the number of participants is quite big. This script offers a quick solution to the challenge of team allocation and participant communication.

## How to 

1. Download and install [Anaconda Navigator](https://anaconda.org/) if you do not have it installed already.

2. Create an [openrouteservice account](https://openrouteservice.org/dev/#/login) if you do not already have one. 

3. Create a [standard token](https://openrouteservice.org/dev/#/home) to use for the API and copy the Key token to your clipboard.

4. Organize a running Dinner and gather the necessary information of the participants. For the script to work you should have two datatables matching the designation **data.xlsx** and **Afterparty.xlsx**. The **data.xlsx** one should contain the fields **"Team Nr."** (for ID), **"Name 1"** and **"Name 2"** (for Mailing), **"Adress"** (for Geocoding), **"will to double"** _([integer](https://www.w3schools.com/python/python_numbers.asp))_ and **"readiness starter"** _([bool](https://www.w3schools.com/python/python_booleans.asp))_ (for assigning Teams the courses they will cook themselves) and **"Ring at"** and **"Allergies or else"** (for the communication of essential information between Hosts and Guests via Mail). At the very least you should have the adresses. If you did not collect data on willingnes to cook twice or inability to cook the first course, we recommend you just create these columns anyway and fill them with homogenous values as the script depends on their presence. The **Afterparty.xlsx** file need only contain the column **"Adress"** containing the Adress of the Afterparty. If you do not have an Afterparty organized, we recommend you either do or type in an Adress located somewhere close to the center of the whole of the participants adresses as the script wont work without an Afterparty Adress and the assignment of courses is done by distance to the Afterparty location. Place both the **data.xlsx** and the **Afterparty.xlsx** files in the same Folder/Directory as the **Full_Script_RDC_V2.IPYNB** file (otherwise the script won't function).

5. Open the **Full_Script_RDC_V2.IPYNB** file with Jupyter Notebook (accesable through Anaconda Navigator), insert the token Key in the first cell, where it says _"""YOUR KEY HERE"""_ (**IMPORTANT NOTE FOR THOSE NOT WELL VERSED IN THE USAGE OF PYTHON**: if your Key is for example ex420, the space in the parenthesis should read _**key= "ex420"**_. There should only be one pair of quotation marks but it is necessary for the code to function, that those will be used ).

6. Run all the cells in their designated order by clicking the "Run \>" Button for each (if a cell has a \[*\] in square brackets next to it, it means, there is still a process of that cell running, so pause the clicking, while it runs the process).

7. If everything went right, there should now be a .txt file for every Team of your running dinner, containing their respective information mail. Also, even if you stop excecuting the script before Mail creation, there should now be the **overview.txt** file in your folder containing the information on who hosts whom for which course.

If this intrigues you but you don't know what a running dinner is, we offer you this [explanation video](https://www.youtube.com/watch?si=_X1kL1hl2W7vYcCO&v=iZEZ5yNHWA8&feature=youtu.be) (so far we only have a german one. Running dinner fans are free to submit links to explanations in other languages).

## Authors

V2: 
* David Benedict <david.benedict@stud.uni-heidelberg.de>

[V1:](https://courses.gistools.geog.uni-heidelberg.de/pk256/running-dinner-calculator)
* David Benedict <david.benedict@stud.uni-heidelberg.de>
* Jakob Tinapp <jakob.tinapp@stud.uni-heidelberg.de>
* Jakob Moser <moser@cl.uni-heidelberg.de>


## Acknowledgments

This project uses the [OpenRouteService](https://openrouteservice.org/) API for route calculations. OpenRouteService is provided by [Heidelberg Institute for Geoinformation Technology (HeiGIT)](https://heigit.org/). 

This project is **not** affiliated with or endorsed by OpenRouteService or HeiGIT. OpenRouteService is a third-party API used in this project to provide routing and geocoding functionalities. Please ensure that your use of this project complies with OpenRouteService's [Terms of Service](https://openrouteservice.org/terms-of-service/) and look into the [API documentation](https://openrouteservice.org/documentation/) for more information.


