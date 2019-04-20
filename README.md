# KJSCE-Writeup-Creator

Welcome to the KJSCE Writeup Creator. Using this script, you can quickly make your writeups without having to manually copy and paste code and most importantly, without having to go online and search for answers for Post Lab Subjective questions. Why? Because our program will do all of that for you! It will get all your code and output screenshots, and paste them in your writeup. Besides, **IT WILL GO ONLINE, SEARCH FOR ANSWERS FOR POSTLAB QUESTIONS, SUMMARISE THEM, AND DIRECTLY PASTE THEM INTO YOUR WRITEUP!**

Amazing, isn't it?

<h2>How to install:</h2>
First, clone this repository onto your computer and run the following command in your cmd (command prompt):

```
pip install -r requirements.txt
```
Or, you could also manually install all dependencies as follows:

Run the following commands in your cmd:
```
pip install python-docx
pip install sumy
```
Run the following command in your git bash:
```
pip install git+https://github.com/abenassi/Google-Search-API
```

<h2>How to use:</h2>

In order to use the program, all you have to do is paste all your code files in the *code* folder and all your screenshots in your *screenshots* folder.

Once that is done, open your write-up and add some text (eg. Code) wherever you want all your code to be pasted and add some other text (eg. Output) wherever you want all your screenshots to be pasted. Also, check the heading under which all the postlab questions lie(eg., Post Lab ). If there is no heading, add something such as "PostLab".

After that, run the writeup.py file in the cmd as follows:
```
python writeup.py
```
Now, the code will ask you the name of the writeup that you want to enter. Once entered, enter the headings that you added for adding the code, output and post lab questions and press enter.

The writeup will now be automatically created and the output will be saved in a file called *output.docx*.

Now, if you want to create another writeup, delete the files in the *code* and *screenshots* folders and add the new code and screenshots and run the writeup.py script as shown before.

<h3>Completing only parts of the writeup</h3>

Follow the steps as listed above and:

1. If you don't want Post Lab Questions to be solved, input *nopostlab* in cmd
2. If you don't want Code to be pasted, input *nocode* in cmd
3. If you don't want Output to be pasted, input *nooutput* in cmd

<h3>Developed by:</h3>

- Sharvai Patil
=======
# KJSCE-Writeup-Creator
