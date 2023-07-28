# # # # # #         Nick Weis's Meal Planning Program         # # # # # # #

# Import necessary modules
import random
import openpyxl
import os
import sys

# Point the OS to the right folder
os.chdir(r'C:\Users\nichow13\OneDrive - kochind.com\Documents\Additional Documents\Python\Mealplanner')

global pantry   # ensures that the variable is global
global stock    # ensures that the variable is global

pantry = openpyxl.load_workbook('PantryDatabase.xlsx')  # Opens the pantry database (excel spreadsheet)
stock = pantry['Sheet1']                                # Opens the correct spreadsheet within the database

# Food Data Initialization
# Current Inventory
    # Ingredients are dictionaries with key/value pairs that describe the nutritional value, portion size, and where
    # to find the current stock within the excel spreadsheet.

global chicken
chicken = {'name': 'Chicken', 'portion': 100, 'units': 'grams',
           'calories': 165, 'protein': 31, 'carbs': 0, 'fat': 3.6,
           'curstock': stock['C3'].value, 'dbid': stock['C3']}

global brocolli
brocolli = {'name': 'Brocolli', 'portion': 100, 'units': 'grams',
            'calories': 35.4, 'protein': 2.4, 'carbs': 7.2, 'fat': 0.4,
            'curstock': stock['C4'].value, 'dbid': stock['C4']}

global mushrooms
mushrooms = {'name': 'Portabella Mushrooms', 'portion': 100, 'units': 'grams',
             'calories': 29, 'protein': 3.3, 'carbs': 4.4, 'fat': 0.6,
             'curstock': stock['C5'].value, 'dbid': stock['C5']}

global eggs
eggs = {'name': 'Eggs', 'portion': 1, 'units': 'eggs',
        'calories': 72, 'protein': 6.3, 'carbs': 0.4, 'fat': 4.8,
        'curstock': stock['C6'].value, 'dbid': stock['C6']}

global bacon
bacon = {'name': 'Bacon', 'portion': 100, 'units': 'grams',
         'calories': 417.9, 'protein': 12.5, 'carbs': 1.4, 'fat': 39.3,
         'curstock': stock['C7'].value, 'dbid': stock['C7']}

global roastbeef
roastbeef = {'name': 'Roast Beef', 'portion': 100, 'units': 'grams',
             'calories': 118.3, 'protein': 18.3, 'carbs': 1.1, 'fat': 3.2,
             'curstock': stock['C8'].value, 'dbid': stock['C8']}

global whitebread
whitebread = {'name': 'White Bread', 'portion': 1, 'units': 'slices',
              'calories': 264.9, 'protein': 8.9, 'carbs': 48.6, 'fat': 3.2,
              'curstock': stock['C9'].value, 'dbid': stock['C9']}

global strawberries
strawberries = {'name': 'Strawberries', 'portion': 100, 'units': 'grams',
                'calories': 32, 'protein': 0.7, 'carbs': 7.7, 'fat': 0.3,
                'curstock': stock['C10'].value, 'dbid': stock['C10']}

global bananas
bananas = {'name': 'Bananas', 'portion': 100, 'units': 'grams',
           'calories': 89, 'protein': 1.1, 'carbs': 22.8, 'fat': 0.3,
           'curstock': stock['C11'].value, 'dbid': stock['C11']}

global orangejuice
orangejuice = {'name': 'Orange Juice', 'portion': 1, 'units': 'cups',
               'calories': 110, 'protein': 2, 'carbs': 25.4, 'fat': 0.3,
               'curstock': stock['C12'].value, 'dbid': stock['C12']}

global twopercentmilk
twopercentmilk = {'name': '2% Milk', 'portion': 1, 'units': 'cups',
                  'calories': 122, 'protein': 8.1, 'carbs': 12.3, 'fat': 4.8,
                  'curstock': stock['C13'].value, 'dbid': stock['C13']}

global salmon
salmon = {'name': 'Salmon', 'portion': 100, 'units': 'grams',
          'calories': 208, 'protein': 20.4, 'carbs': 0, 'fat': 13.4,
          'curstock': stock['C14'].value, 'dbid': stock['C14']}

global whiterice
whiterice = {'name': 'White Rice', 'portion': 1, 'units': 'cups',
             'calories': 685, 'protein': 12.6, 'carbs': 151, 'fat': 1,
             'curstock': stock['C15'].value, 'dbid': stock['C15']}

global oliveoil
oliveoil = {'name': 'Olive Oil', 'portion': 1, 'units': 'tbsp',
            'calories': 119, 'protein': 0, 'carbs': 0, 'fat': 13.5,
            'curstock': stock['C16'].value, 'dbid': stock['C16']}

global yellowonion
yellowonion = {'name': 'Yellow Onion', 'portion': 100, 'units': 'grams',
               'calories': 132, 'protein': 0.9, 'carbs': 7.9, 'fat': 10.8,
               'curstock': stock['C17'].value, 'dbid': stock['C17']}

global greenbellpepper
greenbellpepper = {'name': 'Green Bell Pepper', 'portion': 100, 'units': 'grams',
             'calories': 20, 'protein': 0.9, 'carbs': 4.6, 'fat': 0.2,
             'curstock': stock['C18'].value, 'dbid': stock['C18']}

global tomato
tomato = {'name': 'Tomato', 'portion': 100, 'units': 'grams',
          'calories': 18.0, 'protein': 0.9, 'carbs': 3.9, 'fat': 0.2,
          'curstock': stock['C19'].value, 'dbid': stock['C19']}

global groundbeef90
groundbeef90 = {'name': '90% Lean Ground Beef', 'portion': 100, 'units': 'grams',
                'calories': 788.8, 'protein': 89.6, 'carbs': 0, 'fat': 44.8,
                'curstock': stock['C20'].value, 'dbid': stock['C20']}

global beefbroth
beefbroth = {'name': 'Beef Broth', 'portion': 1, 'units': 'cups',
             'calories': 16.8, 'protein': 2.7, 'carbs': 0.1, 'fat': 0.5,
             'curstock': stock['C21'].value, 'dbid': stock['C21']}

global tomatosauce
tomatosauce = {'name': 'Tomato Sauce', 'portion': 1, 'units': 'cups',
               'calories': 58.8, 'protein': 3.2, 'carbs': 14.1, 'fat': 0.4,
               'curstock': stock['C22'].value, 'dbid': stock['C22']}

global pintobeans
pintobeans = {'name': 'Pinto Beans', 'portion': 1, 'units': 'cups',
              'calories': 247, 'protein': 14, 'carbs': 45, 'fat': 1.7,
              'curstock': stock['C23'].value, 'dbid': stock['C23']}

global blackbeans
blackbeans = {'name': 'Black Beans', 'portion': 1, 'units': 'cups',
             'calories': 241, 'protein': 16, 'carbs': 44, 'fat': 0.8,
             'curstock': stock['C24'].value, 'dbid': stock['C24']}

global chilipowder
chilipowder = {'name': 'Chili Powder', 'portion': 1, 'units': 'tbsp',
               'calories': 23, 'protein': 1.1, 'carbs': 4, 'fat': 1.1,
               'curstock': stock['C25'].value, 'dbid': stock['C25']}

global oregano
oregano = {'name': 'Oregano', 'portion': 1, 'units': 'tbsp',
           'calories': 4.8, 'protein': 0.2, 'carbs': 1.2, 'fat': 0.1,
           'curstock': stock['C26'].value, 'dbid': stock['C26']}

global cumin
cumin = {'name': 'Cumin', 'portion': 1, 'units': 'tsp',
         'calories': 10, 'protein': 0, 'carbs': 1, 'fat': 0,
         'curstock': stock['C27'].value, 'dbid': stock['C27']}

global coriander
coriander = {'name': 'Coriander', 'portion': 1, 'units': 'gram',
             'calories': 3.4, 'protein': 0.2, 'carbs': 0.2, 'fat': 0.2,
             'curstock': stock['C28'].value, 'dbid': stock['C28']}

global salt
salt = {'name': 'Salt', 'portion': 1, 'units': 'tsp',
        'calories': 0, 'protein': 0, 'carbs': 0, 'fat': 0,
        'curstock': stock['C29'].value, 'dbid': stock['C29']}

global cayennepowder
cayennepowder = {'name': 'Cayenne Powder', 'portion': 1, 'units': 'tsp',
                 'calories': 5.7, 'protein': 0.2, 'carbs': 1, 'fat': 0.3,
                 'curstock': stock['C30'].value, 'dbid': stock['C30']}

global garlicpowder
garlicpowder = {'name': 'Garlic Powder', 'portion': 1, 'units': 'tbsp',
                'calories': 32, 'protein': 1.6, 'carbs': 7.1, 'fat': 0.1,
                'curstock': stock['C31'].value, 'dbid': stock['C31']}

global cheddarcheese
cheddarcheese = {'name': 'Cheddar Cheese', 'portion': 100, 'units': 'grams',
                'calories': 403, 'protein': 24.9, 'carbs': 1.3, 'fat': 33.1,
                'curstock': stock['C32'].value, 'dbid': stock['C32']}

# Store ingredients within a list
ingredients = [chicken, brocolli, mushrooms, eggs, bacon, roastbeef, whitebread, twopercentmilk, orangejuice, bananas,
               strawberries, garlicpowder, cayennepowder, salt, coriander, cumin, oregano, chilipowder, blackbeans,
               pintobeans, tomatosauce, tomato, beefbroth, groundbeef90, greenbellpepper, yellowonion, oliveoil,
               whiterice, salmon]

# Create recipes
# Note that servings is how many servings are created by the recipe, curservs is how many are currently in stock.
# Recipes are dictionaries that contain values for

global eggsandbacon
eggsandbacon = {'name': 'Eggs and Bacon',
                'ingrlist': [eggs, bacon],
                'ingrquant': [3, 2],
                'mealtype': 'breakfast',
                'servings': 1,
                'curservs': stock['F3'].value,
                'dbid': stock['F3'],
                'macros': {}}

global eggcups
eggcups = {'name': 'Eggs Cups',
                'ingrlist': [eggs, cheddarcheese],
                'ingrquant': [3, 2],
                'mealtype': 'breakfast',
                'servings': 1,
                'curservs': stock['F3'].value,
                'dbid': stock['F3'],
                'macros': {}}

global chickenstirfry
chickenstirfry = {'name': 'Chicken Stir Fry',
                  'ingrlist': [chicken, brocolli, mushrooms],
                  'ingrquant': [1, 2, 2],
                  'mealtype': 'dinner',
                  'servings': 1,
                  'curservs': stock['F4'].value,
                  'dbid': stock['F4'],
                  'macros': {}}

global roastbeefsandwich
roastbeefsandwich = {'name': 'Roast Beef Sandwich',
                     'ingrlist': [roastbeef, whitebread],
                     'ingrquant': [3, 2],
                     'mealtype': 'lunch',
                     'servings': 1,
                     'curservs': stock['F5'].value,
                     'dbid': stock['F5'],
                     'macros': {}}

global strawberrybananasmoothie
strawberrybananasmoothie = {'name': 'Strawberry-Banana Smoothie',
                            'ingrlist': [strawberries, bananas, orangejuice, twopercentmilk],
                            'ingrquant': [4, 1, 1, 0.5],
                            'mealtype': 'breakfast',
                            'servings': 1,
                            'curservs': stock['F6'].value,
                            'dbid': stock['F6'],
                            'macros': {}}

global salmonandrice
salmonandrice = {'name': 'Salmon and Rice',
                 'ingrlist': [salmon, whiterice],
                 'ingrquant': [1, 1],
                 'mealtype': 'dinner',
                 'servings': 1,
                 'curservs': stock['F7'].value,
                 'dbid': stock['F7'],
                 'macros': {}}

global chili
chili = {'name': 'Chili',
         'ingrlist': [oliveoil, yellowonion, greenbellpepper, garlicpowder, groundbeef90, beefbroth, tomatosauce,
                      tomato, pintobeans, blackbeans, chilipowder, oregano, cumin, coriander, salt, cayennepowder],
         'ingrquant': [1, 1, 1.2, 1, 2, 2, 1, 1, 2, 2, 3, 1, 1, 1, 1, 1],
         'mealtype': 'dinner',
         'servings': 8,
         'curservs': stock['F8'].value,
         'dbid': stock['F8'],
         'macros': {}}

# Store recipes in a list
recipes = [eggsandbacon, chickenstirfry, roastbeefsandwich, strawberrybananasmoothie, salmonandrice, chili]

# Define meal nutrition function:
# Calculates the calories, protein, carbs and fat of a meal
# by adding the macros of the ingredients times the quantity of the ingredients called for by the recipe.

def mealnutrition(meal):

    mealcals = 0
    mealprot = 0
    mealcarbs = 0
    mealfat = 0

    n = 0
    for i in meal['ingrlist']:
        mealcals = mealcals + i['calories'] * meal['ingrquant'][n]
        n = n + 1
    meal['macros']['calories'] = mealcals/meal['servings']

    n = 0
    for i in meal['ingrlist']:
        mealprot = mealprot + i['protein'] * meal['ingrquant'][n]
        n = n + 1
    meal['macros']['protein'] = mealprot/meal['servings']

    n = 0
    for i in meal['ingrlist']:
        mealcarbs = mealcarbs + i['carbs'] * meal['ingrquant'][n]
        n = n + 1
    meal['macros']['carbs'] = mealcarbs/meal['servings']

    n = 0
    for i in meal['ingrlist']:
        mealfat = mealfat + i['fat'] * meal['ingrquant'][n]
        n = n + 1
    meal['macros']['fat'] = mealfat/meal['servings']


for i in recipes:
    if len(i['macros']) == 0:           # Fills in empty meal macros
        mealnutrition(i)

# Create a list of breakfasts
breakfasts = []
for i in recipes:
    if i['mealtype'] == 'breakfast':
        breakfasts.append(i)

# Create a list of lunches
lunches = []
for i in recipes:
    if i['mealtype'] == 'lunch':
        lunches.append(i)

# Create a list of dinners
dinners = []
for i in recipes:
    if i['mealtype'] == 'dinner':
        dinners.append(i)

# Create a list of snacks
snacks = []
for i in recipes:
    if i['mealtype'] == 'snack':
        snacks.append(i)


go = 'go'

print('Hello, welcome to your meal planner\n')

while go == 'go':

    print('What would you like to do?\n')
    print('1. See servings available')
    print('2. See ingredients available')
    print('3. Create a meal calendar')
    print('Exit')
    choice = input('')

# Option 1: See number of servings per meal

    if choice == '1':
        print('Ok, would you like to see all servings or servings for a particular meal?')
        print('1. All servings')
        print('2. A particular meal')
        subchoice = input('')

        if subchoice == '1':
            for i in recipes:
                print(i['name'] + ': ' + str(i['curservs']))
        elif subchoice == '2':
            print('Ok, which meal would you like to see the servings for?')
            for i in recipes:
                print(i['name'])
            subchoice = input('')
            print('Ok, here is the number of servings for ' + subchoice)
            for i in recipes:
                if i['name'] == subchoice:
                    print(str(i['curservs']))

# Option 2: See ingredient stock

    elif choice == '2':
        print('Ok, would you like to see your entire stock or the stock for a particular ingredient?')
        print('1. Entire stock')
        print('2. A particular ingredient')
        subchoice = input('')

        if subchoice == '1':
            for i in ingredients:
                print(i['name'] + ': ' + str(i['curstock']))
        elif subchoice == '2':
            print('Ok, which ingredient would you like to see your stock of?')
            for i in ingredients:
                print(i['name'])
            subchoice = input('')
            print('Ok, here is the number of servings for ' + subchoice)
            for i in ingredients:
                if i['name'] == subchoice:
                    print(str(i['curstock']))

# Option 3: Create meal calendar

    elif choice == '3':
        # Excel Prep
        # Open spreadsheets
        os.chdir(r'C:\Users\nichow13\OneDrive - kochind.com\Documents\Additional Documents\Python\Mealplanner')  # Changes the current directory to the Python folder
        calendar = openpyxl.load_workbook('xltest.xlsx')  # Opens the calendar workbook
        september = calendar['September']  # Opens the calendar sheet

        # Initialize excel iteration variables
        xlcolumns = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
        xlrows = ['5', '7', '9', '11', '13']
        c = 1
        r = 2
        curcol = xlcolumns[c]
        currow = xlrows[r]
        curcell = curcol + currow

        # For loop to choose a daily menu, compare macros, and write to excel calendar for the desired amount of days

        print('How many days would you like to schedule meals for?')
        days = input('')
        for i in range(int(days)):

        # Choose a breakfast, lunch and dinner

            breakfastpool = []  # Meal choosing system prototype 2
            for i in breakfasts:
                if i['curservs'] != 0:  # Adds meals with leftover portions to a list and randomly selects one
                     breakfastpool.append(i)
            if len(breakfastpool) < 3:  # If there is between 0 and 2 meals with leftover portions, a random meal is added to the pool
                surprisebreakfast = random.choice(breakfasts)
                breakfastpool.append(surprisebreakfast)
            else:
                surprisebreakfast = 'none'

            breakfast = random.choice(breakfastpool)
            #print(breakfast['name'])                   #Use to test

            if breakfast == surprisebreakfast:
                # subtract amount of meal ingredients from the current stock
                n = 0
                for i in breakfast['ingrlist']:
                    i['curstock'] = i['curstock'] - breakfast['ingrquant'][n]
                    i['dbid'].value = i['curstock']
                    n = n + 1
            else:
                # subtract a meal serving from the current amount of servings
                breakfast['curservs'] = breakfast['curservs'] - 1
                breakfast['dbid'].value = breakfast['curservs']

            lunchpool = []
            for i in lunches:
                if i['curservs'] != 0:
                    lunchpool.append(i)
            if len(lunchpool) < 3:
                surpriselunch = random.choice(lunches)
                lunchpool.append(surpriselunch)
            else:
                surpriselunch = 'none'

            lunch = random.choice(lunchpool)
            #print(lunch['name'])                       #Use to test

            if lunch == surpriselunch:
                # subtract amount of meal ingredients from the current stock
                n = 0
                for i in lunch['ingrlist']:
                    i['curstock'] = i['curstock'] - lunch['ingrquant'][n]
                    i['dbid'].value = i['curstock']
                    n = n + 1
            else:
                # subtract a meal serving from the current amount of servings
                lunch['curservs'] = lunch['curservs'] - 1
                lunch['dbid'].value = lunch['curservs']

            dinnerpool = []
            for i in dinners:
                if i['curservs'] != 0:
                    dinnerpool.append(i)
            if len(dinnerpool) < 3:
                surprisedinner = random.choice(dinners)
                dinnerpool.append(surprisedinner)
            else:
                surprisedinner = 'none'

            dinner = random.choice(dinnerpool)
            #print(dinner['name'])                      #Use to test

            if dinner == surprisedinner:
                # subtract amount of meal ingredients from the current stock
                n = 0
                for i in dinner['ingrlist']:
                    i['curstock'] = i['curstock'] - dinner['ingrquant'][n]
                    i['dbid'].value = i['curstock']
                    n = n + 1
            else:
                # subtract a meal serving from the current amount of servings
                dinner['curservs'] = dinner['curservs'] - 1
                dinner['dbid'].value = dinner['curservs']

            dailymenu = breakfast['name'] + '\n' + lunch['name'] + '\n' + dinner['name']  # Create a string of the days menu
            september[curcell] = dailymenu  # Write days menu into Excel spreadsheet

            # Updates the cell and row values
            if c == 6:
                c = 0
                r = r + 1
            elif c != 6:
                c = c + 1

            curcol = xlcolumns[c]
            currow = xlrows[r]
            curcell = curcol + currow
            calendar.save('xltest.xlsx')
            pantry.save('PantryDatabase.xlsx')

        os.system("start EXCEL.EXE xltest.xlsx")

        for i in ingredients:
            if i['curstock'] < 0:
                print(str(abs(i['curstock'])) + ' ' +  i['units'] + ' - ' + i['name'])
                shoppinglistfile = open('shoppinglist.txt', 'a')
                shoppinglistfile.write((str(abs(i['curstock']))) + ' ' + i['units'] + ' - ' + i['name'] + '\n')
                shoppinglistfile.close()

# Option Exit:

    elif choice == 'exit' or 'Exit':
        sys.exit('Goodbye!')

# Stop program
go = 'stop'