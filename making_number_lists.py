# aizen_list_of_things_he_likes = ['space', 'sharks', 'dinosaurs', 'humans', 'buildings', 'engineers']
# print (aizen_list_of_things_he_likes)
# popped_item = aizen_list_of_things_he_likes.pop()
# print (popped_item)
# popped_item = popped_item.title()
# print (popped_item)
# aizen_list_of_things_he_likes.append(popped_item)
# print (aizen_list_of_things_he_likes)
# print ()

# print ("I will create a list of numbers and add them together.")
# numbers = list(range(-100,-2 ))
# print (numbers)
# sum_of_numbers = sum(numbers) 
# print ()
# print (f"If we add all the numbers in the list it equals {sum_of_numbers}.")
# print ()
redo = input('Do you want to print a list of numbers? (yes/no)')
while redo == 'yes' or redo == 'y':
	list_of_numbers = []
	lowest_number = int(input('Please enter the lowest number: '))
	highest_number = int(input('Please enter the highest number: '))
	power = int(input('Please enter the power: '))
	values = list(range(lowest_number, highest_number+1))
	for value in values:
		value = (value**power)
		list_of_numbers.append(value)
	print (list_of_numbers)
	print ()
	redo = input('Do you want to print a new list? (yes/no)')
	continue
print ('Press any key to end the program.')
x = input()