Imports System.Data.SqlClient
Imports System.Reflection

' Este archivo contiene funciones "iguales" para cada uno de los accesos a datos más comunmente utilizados.
' Estos accesos a datos se encuentran divididos en clases con el nombre de cada uno, por ejemplo, en la
' clase SQL se podrá acceder a la función Alta que da de alta a registros en una tabla de una DB de SQL.
' La función Alta de la clase OLEDB hace los mismo con tablas de DB OLEDB.

' Otra forma de organizar lógicamente sería separar en diferentes archivos cada una de las clases pero referen-
' ciarlas dentro de un mismo Namesapce.

' Todas las funciones y procedimiento de la clase SQL son exclusivamente para DB en SQL Server.

' Para la creación de las clases con la estructura de los datos pueden bajar un generador de clases que publiqué
' también el la web del programador o solicitármelo a visualman2001@hotmail.com (este generador crea clases con
' la estructura de una tabla de una base de datos y la graba en un archivo VB, con lo cual con dicho generador y
' estas funciones se puede tener un ABM sencillo en muy poco tiempo, dejándolo solamente la tearea de crear los
' formularios (o páginas web, ya que funciona de la misma manera por se programación en capas).

Public Class Sql
    ' Esta función recibe una cadena con la conexión a la DB, el nombre de una tabla, un filtro (preferiblemente con una clave única) y un objeto definido por
    ' el usuario y carga los datos del registro recuperado en el objeto (este es recibido por referencia y no por valor). Si se recupera más de un registro se
    ' toma el primero y se ignora el resto, si no se recupera ningún registro, el objeto retorna con los mismos valores con los que se recibió.
    '
    ' El objeto tiene que tener como mínimo tantas propiedades como campos la estructura de datos (si tuviera más, estas serán obviadas ya que sólo se toman
    ' los campos que hay en la difinición de la tabla, recorriendo su estructura).
    '
    ' Aunque permite recibir una lista de campos (o propiedades) que no deben ser llenadas, la función no limpia el contenido de dichos campos (o propiedades), o sea
    ' que si las mismas tenían algún tipo de valor, éste no es removido, sinó dejado tal como están.
    ' 
    ' Ej. llámada típica: Cargar("< conexión >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla")  -- Acá se carga el objeto obviando el campo IDtabla
    ' Ej. llamada más de una excepción: Cargar("< conexión >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla", "imagen")  -- Acá se carga el objeto obviando los campos IDtabla e imagen
    ' Ej. llámada sin excepción: Cargar("< conexión >", "< tabla >", "< filtro >", < variable objeto >, "")  -- Acá se carga el objeto pero no se obvia ningún campo (en este caso el parámetro es opcional)
    '
    ' Si la función no encuentra problemas se retorna una cadena vacía, sinó se retorna un mensaje conteniendo el error ocurrido.

    Public Shared Function Cargar(ByVal conexion As String, ByVal tabla As String, ByVal filtro As String, ByRef objeto As Object, ByVal ParamArray excepcion() As String) As String
        ' Evaluo el filtro
        If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio después del WHERE es parte de la cadena
            filtro = "WHERE " & filtro
        End If

        ' Defino variables de acceso y manipulación de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla & " " & filtro, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn

        ' Defino variables de propósitos generales
        Dim ex As String
        Dim ok As Boolean
        Dim tp As Type = objeto.GetType() ' guardo el tipo de objeto recibido por "reflexión"

        ' Cargo la tabla "en memoria" (el nombre "tabla" es sólo para la tabla virtual)
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

            ' Retorno que se logro grabar el registro (cadena vacía = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta función recibe una cadena con la conexión a la DB, el nombre de la tabla donde se quiere agregar un registro y un objeto definido por el usuario.
    ' El objeto tiene que tener como mínimo tantas propiedades como campos la estructura de datos (si tuviera más, estas serán obviadas ya que sólo se toman
    ' los campos que hay en la difinición de la tabla, recorriendo su estructura).
    '
    ' Además se le puede pasar un campo de excepión que contiene uno o más campos a exceptuar durante el alta del registro. Esto puede ser útil si hay al-
    ' gunos campos que deban ser obviados durante el alta, como por ejemplo un campo auto-numérico o algún campo que no deba ser inicializado durante su
    ' alta sinó luego sólo por modificaciones, como resultado de acciones, etc.
    ' 
    ' Ej. llámada típica: Alta("< conexión >", "< tabla >", < variable objeto >, "IDtabla")  -- Acá se obvia el campo IDtabla
    ' Ej. llamada más de una excepción: Alta("< conexión >", "< tabla >", < variable objeto >, "IDtabla", "imagen")  -- Acá se obvia el campo IDtabla y el cmapo imagen
    ' Ej. llámada sin excepción: Alta("< conexión >", "< tabla >", < variable objeto >, "")  -- Acá no se obvia ningún campo (en este caso el parámetro es opcional)
    '
    ' Hay que reemplazar "< conexión >" y "< tabla >" por las cadenas correspondientes, mientras que < variable objeto > es el nombre de la variable de objeto.
    ' Esta función retorna una cadena vacía si pudo realizar exitosamente el alta o sino devuelve una cadena conteniendo una descripción del error.

    Public Shared Function Alta(ByVal conexion As String, ByVal tabla As String, ByVal objeto As Object, ByVal ParamArray excepcion() As String) As String
        ' Defino variables de acceso y manipulación de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn
        Dim dr As DataRow

        ' Defino variables de propósitos generales
        Dim ex As String
        Dim ok As Boolean
        Dim tp As Type = objeto.GetType() ' guardo el tipo de objeto recibido por "reflexión"

        ' Cargo la tabla "en memoria" (el nombre "tabla" es sólo para la tabla virtual)
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

            ' Confirmo los cambios realizados y los reflejo en la tabla "física"
            da.Update(ds.Tables("tabla"))
            ds.AcceptChanges()

            ' Retorno que se logro grabar el registro (cadena vacía = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta función recibe una cadena con la conexión a la DB, el nombre de una tabla, un filtro y un objeto definido por el usuario y actualiza los datos del
    ' registro en la DB con los datos del objeto.
    '
    ' El objeto tiene que tener como mínimo tantas propiedades como campos la estructura de datos (si tuviera más, estas serán obviadas ya que sólo se toman
    ' los campos que hay en la difinición de la tabla, recorriendo su estructura).
    '
    ' Además se le puede pasar un campo de excepión que contiene uno o más campos a exceptuar durante el alta del registro. Esto puede ser útil si hay al-
    ' gunos campos que deban ser obviados durante la edición, como por ejemplo un campo auto-numérico o algún campo que no deba ser modificado directamente
    ' sinó como resultado de acciones, etc.
    ' 
    ' Ej. llámada típica: Editar("< conexión >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla")  -- Acá se obvia el campo IDtabla
    ' Ej. llamada más de una excepción: Editar("< conexión >", "< tabla >", "< filtro >", < variable objeto >, "IDtabla", "imagen")  -- Acá se obvia el campo IDtabla y el cmapo imagen
    ' Ej. llámada sin excepción: Editar("< conexión >", "< tabla >", "< filtro >", < variable objeto >, "")  -- Acá no se obvia ningún campo (en este caso el parámetro es opcional)
    '
    ' Hay que reemplazar "< conexión >" y "< tabla >" por las cadenas correspondientes, mientras que < variable objeto > es el nombre de la variable de objeto.
    ' Esta función retorna una cadena vacía si pudo realizar exitosamente la modificación o sino devuelve una cadena conteniendo una descripción del error.

    Public Shared Function Editar(ByVal conexion As String, ByVal tabla As String, ByVal filtro As String, ByVal objeto As Object, ByVal ParamArray excepcion() As String) As String
        ' Evaluo el filtro
        If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio después del WHERE es parte de la cadena
            filtro = "WHERE " & filtro
        End If

        ' Defino variables de acceso y manipulación de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla & " " & filtro, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn
        Dim dr As DataRow

        ' Defino variables de propósitos generales
        Dim ex As String
        Dim ok As Boolean
        Dim tp As Type = objeto.GetType() ' guardo el tipo de objeto recibido por "reflexión"

        ' Cargo la tabla "en memoria" (el nombre "tabla" es sólo para la tabla virtual)
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

                ' Recorro la lista de campos a exceptuar en la edición (puede no haber ninguno)
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

            ' Confirmo los cambios realizados y los reflejo en la tabla "física"
            da.Update(ds.Tables("tabla"))
            ds.AcceptChanges()

            ' Retorno que se logro grabar el registro (cadena vacía = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta función recibe una cadena con la conexión a la DB, el nombre de una tabla, un filtro y elimina los registro recuperados.
    ' Si el filtro llegase a recuperar más de un registro, la función los eliminará a todos.
    ' En el caso de no haber ningún registro para eliminar, no se generará un error, en su lugar se retornará un cadena indicando
    ' que no se recuperaron registros.
    '
    ' Ej. llámada típica: Quitar("< conexión >", "< tabla >", "< filtro >")
    '
    ' Si la función se ejecuta sin errores se devuelve una cadena vacía, sino se devuelve una cadena conteniendo el error ocurrido.

    Public Shared Function Quitar(ByVal conexion As String, ByVal tabla As String, ByVal filtro As String) As String
        ' Evaluo el filtro
        If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio después del WHERE es parte de la cadena
            filtro = "WHERE " & filtro
        End If

        ' Defino variables de acceso y manipulación de datos
        Dim cn As New SqlConnection(conexion)
        Dim da As New SqlDataAdapter("SELECT * FROM " & tabla & " " & filtro, cn)
        Dim cb As New SqlCommandBuilder(da)
        Dim ds As New DataSet

        Dim dc As DataColumn
        Dim dr As DataRow

        ' Cargo la tabla "en memoria" (el nombre "tabla" es sólo para la tabla virtual)
        da.Fill(ds, "tabla")

        ' Verifico la cantidad de filas
        If ds.Tables("tabla").Rows.Count = 0 Then
            Return "No se ha encontrado ningún registro que cumpla con la condición solicitada."
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

            ' Confirmo los cambios realizados y los reflejo en la tabla "física"
            da.Update(ds.Tables("tabla"))
            ds.AcceptChanges()

            ' Retorno que se logro grabar el registro (cadena vacía = no error)
            Return ""

        Catch en As Exception
            ' Retorno que no se logro grabar el registro (indicando mensaje de error)
            Return en.Message.ToString
        End Try
    End Function

    ' Esta función recibe una lista de parámetros y carga y devuelve un Dataset.
    ' Se le puede también añadir filtro y orden (ambos opcionales).

    Public Shared Function Llenar(ByVal conexion As String, ByVal origen As String, Optional ByVal filtro As String = "", Optional ByVal orden As String = "") As DataTable
        ' Evaluo el filtro
        If filtro <> "" Then
            If Left(filtro, 6).ToUpper <> "WHERE " Then ' el espacio después del WHERE es parte de la cadena
                filtro = "WHERE " & filtro
            End If
        End If

        ' Evaluo el orden
        If orden <> "" Then
            If Left(orden, 9).ToUpper <> "ORDER BY " Then ' el espacio después del WHERE es parte de la cadena
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
