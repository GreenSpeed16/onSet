# Imports
import matplotlib.pyplot as plt

# Get values for all grades
grade_x_axis = []
grade_amount = 0
grade_list = []
for i in range(0,10):
    grade_amount = int(input('How many V' + str(i) + "'s do you want? "))
    grade_list.append(grade_amount)
    grade_x_axis.append('V' + str(i))

# Create a graph based on the grades
plt.bar(grade_x_axis, grade_list)
plt.show()
