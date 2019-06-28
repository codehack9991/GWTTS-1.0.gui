# GWTTS-1.0.gui - Global Working Time Tracking System for an institution

(Note : This project was done by me as a part of an industrial internship.) 
Aim of this project was to make a desktop application which will take the working time data of all the employees from their access cards' entry information (in form of either excel sheet or csv file), and calculate various required features (like for example- Total working time in a week of the employees, find the laaging employees, calculation of their salaries etc.)

This code was originally written in Jupyter notebook, but for the sake of converting it into a desktop application which even a non-technical employee can also use, we converted it into .py file and then was converted into an deployable format. 

Various methods for deployment were also used, so as to cater to all the different scenarios. They are as follows :

## Methods of Deployment:

      - 1. Converting into an executable format (i.e .exe)
      - 2. Using jython dependency and run this code via java IDE 
            This was done keeping in mind the status of current industrial scenario, where most of the people are still working on Java or related environment. Hence deploying in this format was decided, but ultimately, this method was faced with variety of difficulties. Some of them include :
            - A. Python libraries are generally Cython based, i.e baseed on C/C++. So many libraaries face problems to run using Jython interpreter.
            - B. We tried to find other interpreters, but generally most of them are not stand alone in nature. They ought to require atleast a python interpreter in the background on the host system to run/execute.



