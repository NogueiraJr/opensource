def bucketsort( A ):
  # get hash codes
  code = hashing( A )
  buckets = [list() for _ in range( code[1] )]
  # distribute data into buckets: O(n)
  for i in A:
    x = re_hashing( i, code )
    buck = buckets[x]
    buck.append( i )

  for bucket in buckets:
    insertionsort( bucket )
    ndx = 0
  # merge the buckets: O(n)
  for b in range( len( buckets ) ):
    for v in buckets[b]:
      A[ndx] = v
      ndx += 1
 
import math
 
def hashing( A ):
  m = A[0]
  for i in range( 1, len( A ) ):
    if ( m < A[i] ):
      m = A[i]
  result = [m, int( math.sqrt( len( A ) ) )]
  return result
 
 
def re_hashing( i, code ):
  return int( i / code[0] * ( code[1] - 1 ) )

def insertionsort( aList ):
  for i in range( 1, len( aList ) ):
    tmp = aList[i]
    k = i 
    while k > 0 and tmp < aList[k - 1]:
        aList[k] = aList[k - 1]
        k -= 1
    aList[k] = tmp


count = int(input("Informe quantidade: "))
sort_list = []
for i in range(count):
    sort_list.append(int(input("Informe numero: ")))

bucketsort(sort_list)

print(sort_list)
