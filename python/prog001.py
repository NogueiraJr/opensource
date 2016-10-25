# empty dictionary
d = {}
d['dog'] = 1
d['cat'] = 2
 
# iterating over (key, value) pairs
for key, value in d.iteritems():
  print key, value
   
# dictionaries from tuple list
d1 = dict([('tiger', 1), ('lion', 2)])
d2 = dict(zip(['dog', 'cat'], [1, 2]))
 
# iterating over keys
for key in d1:
  print key, d1[key]
 
# dictionaries with two keys
d1 = {'dog': 1, 'cat': 2}
d2 = dict(wolf=1, snake=2)
 
# iterating over keys
for key in d2:
  print key, d2[key]
