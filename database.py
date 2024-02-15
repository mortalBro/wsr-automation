import mysql.connector

def activeUserMirrorCalculation():
  connection = mysql.connector.connect(
      host='localhost',
      user='bhaiji',
      password='triazine@123',
      database='mortal'
  )
  cursor = connection.cursor()
  query = '''SELECT count(*)
  FROM mortal.master_user
  where status =1;'''

  cursor.execute(query)
  results = cursor.fetchall()
  integer_value = results[0][0]
  return integer_value
