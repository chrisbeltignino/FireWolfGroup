Imports System.Data.SqlClient
Imports System.Reflection

' Este archivo contiene funciones "iguales" para cada uno de los accesos a datos m�s comunmente utilizados.
' Estos accesos a datos se encuentran divididos en clases con el nombre de cada uno, por ejemplo, en la
' clase SQL se podr� acceder a la funci�n Alta que da de alta a registros en una tabla de una DB de SQL.
' La funci�n Alta de la clase OLEDB hace los mismo con tablas de DB OLEDB.

' Otra forma de organizar l�gicamente ser�a separar en diferentes archivos cada una de las clases pero referen-
' ciarlas dentro de un mismo Namesapce.

' Todas las funciones y procedimiento de la clase SQL son exclusivamente para DB en SQL Server.

' Para la creaci�n de las clases con la estructura de los datos pueden bajar un generador de clases que publiqu�
' tambi�n el la web del programador o solicit�rmelo a visualman2001@hotmail.com (este generador crea clases con
' la estructura de una tabla de una base de datos y la graba en un archivo VB, con lo cual con dicho generador y
' estas funciones se puede tener un ABM sencillo en muy poco tiempo, dej�ndolo solamente la tearea de crear los
' formularios (o p�ginas web, ya que funciona de la misma manera por se programaci�n en capas).

Public Class Sql
    ' Esta funci�n recibe una cadena con la conexi�n a la DB, el nombre de una tabla, un filtro (preferiblemente con una clave �nica) y un objeto definido por
    ' el usuario y carga los datos del registro recuperado en el objeto (este es recibido por referencia y no por valor). Si se recupera m�s de un registro se
    ' toma el primero y se ignora el resto, si no se recupera ning�n registro, el objeto retorna con los mismos valores con los que se recibi�.
    '
    ' El objeto tiene que tener como m�nimo tantas propiedades como campos la estructura de datos (si tuviera m�s, estas ser�n obviadas ya que s�lo se toman
    ' los campos que hay en la difinici�n de la tabla, recorriendo su estructura).
    '
    ' Aunque permite recibir una lista de campos (o propiedades) que no deben ser llenadas, la funci�n no limpia el contenido de dichos campos (o propiedades), o sea
    ' que si las mismas ten�an alg�n tipo de valor, �ste no es removido, sin� dejado tal como est�n.
    ' 
    ' Ej. ll�mada t�pica: Cargar("< conexi�n >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla")  -- Ac� se carga el objeto obviando el campo IDtabla
    ' Ej. llamada m�s de una excepci�n: Cargar("< conexi�n >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla", "imagen")  -- Ac� se carga el objeto obviando los campos IDtabla e imagen
    ' Ej. ll�mada sin excepci�n: Cargar("< conexi�n >", "< tabla >", "< filtro >", < variable objeto >, "")  -- Ac� se carga el objeto pero no se obvia ning�n campo (en este caso el par�metro es opcional)
    '
    ' Si la funci�n no encuentra problemas se retorna una cadena vac�a, sin� se retorna un mensaje conteniendo el error ocurrido.

    Public Shared Function Cargar(ByVal conexion As String, ByVal tabla As String, ByVal filtro As String, ByRef objeto As Object, ByVal ParamArray excepcion() As String) As String
        ' Evaluo el filtro
        If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio despu�s del WHERE es parte de la cadena
            filtro = "WHERE " & filtro
        End If

        ' Defino variables de acceso y manipulaci�n de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla & " " & filtro, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn

        ' Defino variables de prop�sitos generales
        Dim ex As String
        Dim ok As Boolean
        Dim tp As Type = objeto.GetType() ' guardo el tipo de objeto recibido por "reflexi�n"

        ' Cargo la tabla "en memoria" (el nombre "tabla" es s�lo para la tabla virtual)
        da.Fill(ds, "tabla")

        ' Verifico la cantidad de filas
        If ds.Tables("tabla").Rows.Count = 0 Then
            Return ""
        End If

        ' Manejador de excepciones (errores)
        Try
            ' Recorro la lista de campos en el registro "virtual"
            For Each dc In ds.Tables("tabla").Columns()
                ' Inicializo variables
                ok = True ' por defecto asumo que el cmapo no debe exceptuarse

                ' Recorro la lista de campos a exceptuar en la carga (puede no haber ninguno)
                For Each ex In excepcion
                    If dc.ColumnName.ToString.ToLower = ex.ToString.ToLower Then
                        ok = False ' indica que este campo en particular debe exceptuarse
                    End If
                Next

                ' Verifico el resultado de recorrer el array de excepciones
                If ok Then
                    ' Defino una variable para la propiedad actual
                    Dim pr As PropertyInfo = tp.GetProperty(dc.ColumnName)

                    ' Cargo la propiedad del objeto con el cmapo del registro (tras verificar el tipo de dato)
                    If pr.PropertyType.Name.ToLower = "string" Then
                        pr.SetValue(objeto, ds.Tables("tabla").Rows(0)(dc.ColumnName).ToString(), Nothing)
                    Else
                        pr.SetValue(objeto, ds.Tables("tabla").Rows(0)(dc.ColumnName), Nothing)
                    End If
                End If
            Next

            ' Retorno que se logro grabar el registro (cadena vac�a = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta funci�n recibe una cadena con la conexi�n a la DB, el nombre de la tabla donde se quiere agregar un registro y un objeto definido por el usuario.
    ' El objeto tiene que tener como m�nimo tantas propiedades como campos la estructura de datos (si tuviera m�s, estas ser�n obviadas ya que s�lo se toman
    ' los campos que hay en la difinici�n de la tabla, recorriendo su estructura).
    '
    ' Adem�s se le puede pasar un campo de excepi�n que contiene uno o m�s campos a exceptuar durante el alta del registro. Esto puede ser �til si hay al-
    ' gunos campos que deban ser obviados durante el alta, como por ejemplo un campo auto-num�rico o alg�n campo que no deba ser inicializado durante su
    ' alta sin� luego s�lo por modificaciones, como resultado de acciones, etc.
    ' 
    ' Ej. ll�mada t�pica: Alta("< conexi�n >", "< tabla >", < variable objeto >, "IDtabla")  -- Ac� se obvia el campo IDtabla
    ' Ej. llamada m�s de una excepci�n: Alta("< conexi�n >", "< tabla >", < variable objeto >, "IDtabla", "imagen")  -- Ac� se obvia el campo IDtabla y el cmapo imagen
    ' Ej. ll�mada sin excepci�n: Alta("< conexi�n >", "< tabla >", < variable objeto >, "")  -- Ac� no se obvia ning�n campo (en este caso el par�metro es opcional)
    '
    ' Hay que reemplazar "< conexi�n >" y "< tabla >" por las cadenas correspondientes, mientras que < variable objeto > es el nombre de la variable de objeto.
    ' Esta funci�n retorna una cadena vac�a si pudo realizar exitosamente el alta o sino devuelve una cadena conteniendo una descripci�n del error.

    Public Shared Function Alta(ByVal conexion As String, ByVal tabla As String, ByVal objeto As Object, ByVal ParamArray excepcion() As String) As String
        ' Defino variables de acceso y manipulaci�n de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn
        Dim dr As DataRow

        ' Defino variables de prop�sitos generales
        Dim ex As String
        Dim ok As Boolean
        Dim tp As Type = objeto.GetType() ' guardo el tipo de objeto recibido por "reflexi�n"

        ' Cargo la tabla "en memoria" (el nombre "tabla" es s�lo para la tabla virtual)
        da.Fill(ds, "tabla")

        ' Manejador de excepciones (errores)
        Try
            ' Inicializo una variable con la estructura del registro
            dr = ds.Tables("tabla").NewRow

            ' Recorro la lista de campos en el registro "virtual"
            For Each dc In ds.Tables("tabla").Columns()
                ' Inicializo variables
                ok = True ' por defecto asumo que el cmapo no debe exceptuarse

                ' Recorro la lista de campos a exceptuar en el alta (puede no haber ninguno)
                For Each ex In excepcion
                    If dc.ColumnName.ToString.ToLower = ex.ToString.ToLower Then
                        ok = False ' indica que este campo en particular debe exceptuarse
                    End If
                Next

                ' Verifico el resultado de recorrer el array de excepciones
                If ok Then
                    ' Defino una variable para la propiedad actual
                    Dim pr As PropertyInfo = tp.GetProperty(dc.ColumnName)

                    ' Cargo el campo del registro con el contenido de la propiedad del objeto
                    dr(dc.ColumnName) = pr.GetValue(objeto, Nothing)
                End If
            Next

            ' Agrego el registro a la tabla "virtual"
            ds.Tables("tabla").Rows.Add(dr)

            ' Confirmo los cambios realizados y los reflejo en la tabla "f�sica"
            da.Update(ds.Tables("tabla"))
            ds.AcceptChanges()

            ' Retorno que se logro grabar el registro (cadena vac�a = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta funci�n recibe una cadena con la conexi�n a la DB, el nombre de una tabla, un filtro y un objeto definido por el usuario y actualiza los datos del
    ' registro en la DB con los datos del objeto.
    '
    ' El objeto tiene que tener como m�nimo tantas propiedades como campos la estructura de datos (si tuviera m�s, estas ser�n obviadas ya que s�lo se toman
    ' los campos que hay en la difinici�n de la tabla, recorriendo su estructura).
    '
    ' Adem�s se le puede pasar un campo de excepi�n que contiene uno o m�s campos a exceptuar durante el alta del registro. Esto puede ser �til si hay al-
    ' gunos campos que deban ser obviados durante la edici�n, como por ejemplo un campo auto-num�rico o alg�n campo que no deba ser modificado directamente
    ' sin� como resultado de acciones, etc.
    ' 
    ' Ej. ll�mada t�pica: Editar("< conexi�n >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla")  -- Ac� se obvia el campo IDtabla
    ' Ej. llamada m�s de una excepci�n: Editar("< conexi�n >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla", "imagen")  -- Ac� se obvia el campo IDtabla y el cmapo imagen
    ' Ej. ll�mada sin excepci�n: Editar("< conexi�n >", "< tabla >", "< filtro >", < variable objeto >, "")  -- Ac� no se obvia ning�n campo (en este caso el par�metro es opcional)
    '
    ' Hay que reemplazar "< conexi�n >" y "< tabla >" por las cadenas correspondientes, mientras que < variable objeto > es el nombre de la variable de objeto.
    ' Esta funci�n retorna una cadena vac�a si pudo realizar exitosamente la modificaci�n o sino devuelve una cadena conteniendo una descripci�n del error.

    Public Shared Function Editar(ByVal conexion As String, ByVal tabla As String, ByVal filtro As String, ByVal objeto As Object, ByVal ParamArray excepcion() As String) As String
        ' Evaluo el filtro
        If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio despu�s del WHERE es parte de la cadena
            filtro = "WHERE " & filtro
        End If

        ' Defino variables de acceso y manipulaci�n de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla & " " & filtro, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn
        Dim dr As DataRow

        ' Defino variables de prop�sitos generales
        Dim ex As String
        Dim ok As Boolean
        Dim tp As Type = objeto.GetType() ' guardo el tipo de objeto recibido por "reflexi�n"

        ' Cargo la tabla "en memoria" (el nombre "tabla" es s�lo para la tabla virtual)
        da.Fill(ds, "tabla")

        ' Verifico la cantidad de filas
        If ds.Tables("tabla").Rows.Count = 0 Then
            Return ""
        End If

        ' Cargo el registro en una variables
        dr = ds.Tables("tabla").Rows(0)

        ' Manejador de excepciones (errores)
        Try
            ' Recorro la lista de campos en el registro "virtual"
            For Each dc In ds.Tables("tabla").Columns()
                ' Inicializo variables
                ok = True ' por defecto asumo que el cmapo no debe exceptuarse

                ' Recorro la lista de campos a exceptuar en la edici�n (puede no haber ninguno)
                For Each ex In excepcion
                    If dc.ColumnName.ToString.ToLower = ex.ToString.ToLower Then
                        ok = False ' indica que este campo en particular debe exceptuarse
                    End If
                Next

                ' Verifico el resultado de recorrer el array de excepciones
                If ok Then
                    ' Defino una variable para la propiedad actual
                    Dim pr As PropertyInfo = tp.GetProperty(dc.ColumnName)

                    ' Cargo el campo del registro con el contenido de la propiedad del objeto
                    dr(dc.ColumnName) = pr.GetValue(objeto, Nothing)
                End If
            Next

            ' Confirmo los cambios realizados y los reflejo en la tabla "f�sica"
            da.Update(ds.Tables("tabla"))
            ds.AcceptChanges()

            ' Retorno que se logro grabar el registro (cadena vac�a = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta funci�n recibe una cadena con la conexi�n a la DB, el nombre de una tabla, un filtro y elimina los registro recuperados.
    ' Si el filtro llegase a recuperar m�s de un registro, la funci�n los eliminar� a todos.
    ' En el caso de no haber ning�n registro para eliminar, no se generar� un error, en su lugar se retornar� un cadena indicando
    ' que no se recuperaron registros.
    '
    ' Ej. ll�mada t�pica: Quitar("< conexi�n >", "< tabla >", "< filtro >")
    '
    ' Si la funci�n se ejecuta sin errores se devuelve una cadena vac�a, sino se devuelve una cadena conteniendo el error ocurrido.

    Public Shared Function Quitar(ByVal conexion As String, ByVal tabla As String, ByVal filtro As String) As String
        ' Evaluo el filtro
        If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio despu�s del WHERE es parte de la cadena
            filtro = "WHERE " & filtro
        End If

        ' Defino variables de acceso y manipulaci�n de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla & " " & filtro, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn
        Dim dr As DataRow

        ' Cargo la tabla "en memoria" (el nombre "tabla" es s�lo para la tabla virtual)
        da.Fill(ds, "tabla")

        ' Verifico la cantidad de filas
        If ds.Tables("tabla").Rows.Count = 0 Then
            Return "No se ha encontrado ning�n registro que cumpla con la condici�n solicitada."
        End If

        ' Cargo el registro en una variables
        dr = ds.Tables("tabla").Rows(0)

        ' Manejador de excepciones (errores)
        Try
            ' Comienzo a eliminar registros
            For Each dr In ds.Tables("tabla").Rows
                ' Elimino el registro
                dr.Delete()
            Next

            ' Confirmo los cambios realizados y los reflejo en la tabla "f�sica"
            da.Update(ds.Tables("tabla"))
            ds.AcceptChanges()

            ' Retorno que se logro grabar el registro (cadena vac�a = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta funci�n recibe una lista de par�metros y carga y devuelve un Dataset.
    ' Se le puede tambi�n a�adir filtro y orden (ambos opcionales).

    Public Shared Function Llenar(ByVal conexion As String, ByVal origen As String, Optional ByVal filtro As String = "", Optional ByVal orden As String = "") As DataTable
        ' Evaluo el filtro
        If filtro <> "" Then
            If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio despu�s del WHERE es parte de la cadena
                filtro = "WHERE " & filtro
            End If
        End If

        ' Evaluo el orden
        If orden <> "" Then
            If Left(orden, 9).ToUpper <> "ORDER BY " Then ' el espacio despu�s del WHERE es parte de la cadena
                orden = "ORDER BY " & orden
            End If
        End If

        ' Defino variables
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter(origen & " " & filtro & " " & orden, cn)
        Dim ds As New DataSet

        ' Cargo la tabla virtual
        da.Fill(ds, "tabla")

        ' Retorno el dataset cargado
        Return ds.Tables("tabla")
    End Function
End Class
