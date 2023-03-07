#LIMPIAMOS SHELL
Clear-Host


#IMPORTAMOS EL MODULO DE ACTIVE DIRECTORY
Import-Module ActiveDirectory


# OBTENER INFORMACIÓN DEL DOMINIO
Write-Host "DATOS DEL DOMINIO AL QUE ESTÁ CONECTADO:" -ForegroundColor Yellow
$Equipo = Get-ADDomainController -Filter *
$Dominio = Get-ADDomain
Write-Host "Dominio:" $Dominio -ForegroundColor Green 
Write-Host "Equipo:" $Equipo -ForegroundColor Green 


#DEFINICION DE FUNCION EXPORTAR USUARIOS Y GRUPOS DEL DOMINIO A .CSV
function ExportaraCSV {
  try {
    #Ruta donde se va a almacenar el .CSV con los datos
    $RutaCSV = "C:\Users\Administrador\Desktop\Datos.CSV"


    #Define las propiedades de usuario que se exportarán al archivo CSV
    $Propiedades = "SamAccountName", "Name", "EmailAddress", "Title", "Department", "Company"

    #Realiza la búsqueda de usuarios y exporta sus propiedades a un archivo CSV
    Get-ADUser -Filter * -SearchBase $Dominio -Server $Equipo -Properties $Propiedades | Select-Object $Propiedades | Export-CSV -Path $RutaCSV -NoTypeInformation
    Write-Host "LOS USUARIOS SE EXPORTARON CORRECTAMENTE..." -ForegroundColor Green
    Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host
  }
  catch {
    Write-Host "HUBO UN PROBLEMA Y NO SE EXPORTARON LOS USUARIOS..." -ForegroundColor Red
  }
  
}
  

#DEFINICIÓN DE FUNCIÓN PARA LA MODIFICACIÓN DE ALGUNAS PROPIEDADES DE ALGUNOS USUARIOS A PARTIR DE UN .CSV
Function CambiarPropiedadesUsuariosCSV {
  # Ruta del archivo .csv a importar
  $RutaPropiedadesCSV = "C:\Users\Administrador\Desktop\PropiedadesNuevas.csv"
  
  #Obtener valores del .csv
  $usuarios = Import-Csv -Path $RutaPropiedadesCSV

  #Bucle para ir modificando los datos de los usuarios a partir del .csv
  foreach ($usuario in $usuarios) {
    Set-ADUser -Identity $usuario.SAMAccountName -Title $usuario.Title -Company $usuario.Company -Department $usuario.Department -EmailAddress $usuario.EmailAddress
  }
  Write-Host "LOS ATRIBUTOS PARA LOS USUARIOS SE MODIFICARON CORRECTAMENTE..." -ForegroundColor Green
  Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host
}


# DEFINICIÓN DE FUNCIÓN PARA CREAR Y ADMINISTRAR DIRECTIVAS DE GRUPO
function GestionarGPO {
  #Variable para el nombre de la GPO 
  $NombreGPO = Read-Host "Inserte un nombre para la GPO a crear"

  #Comprobamos si la GPO ya existe, si no, la creamos
  if (!(Get-GPO -Name $NombreGPO -ErrorAction SilentlyContinue)) {
    New-GPO -Name $NombreGPO 
    Write-Output "La GPO $NombreGPO ha sido creada"
  }
  else {
    Write-Output "La GPO $NombreGPO ya existe"
  }

  #Configuramos la configuración de la GPO
  Set-GPRegistryValue -Name $NombreGPO -Key "HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -ValueName "NoAutoUpdate" -Type DWORD -Value 1
  Set-GPRegistryValue -Name $NombreGPO -Key "HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -ValueName "ScheduledInstallDay" -Type DWORD -Value 0
  Set-GPRegistryValue -Name $NombreGPO -Key "HKLM\Software\Policies\Microsoft\Windows\WindowsUpdate\AU" -ValueName "ScheduledInstallTime" -Type DWORD -Value 3

  $opcion = Read-Host "¿Desea crear una nueva unidad organizativa? (si/no)"

  if ($opcion -eq "si") {
    Write-Host "Vamos a crear una nueva unidad organizativa..."
    #Condicional para crear una nueva unidad organizativa
    $OU = Read-Host "Introduzca nombre de unidad organizativa a crear"

    #Comprobar si la unidad organizativa existe
    $ExistenciaOU = Get-ADOrganizationalUnit -Filter "Name -eq '$OU'" -SearchBase $Dominio -ErrorAction SilentlyContinue

    #Si no existe la unidad organizativa, la creamos
    if (!$ExistenciaOU) {
      New-ADOrganizationalUnit -Name $OU -Path $Dominio
      Write-Output "Se ha creado la nueva OU '$OU'."
      #Asignamos la GPO a una unidad organizativa (OU)
      $NombreOU = Read-Host "Inserte el nombre de unidad organizativa donde se aplicará la GPO"
      $ruta = 'OU=' + $NombreOU + ',' + $Dominio
      $RutaGPC = "CN={$(Get-GPO -Name $NombreGPO).Id},CN=Policies,CN=System,$($ruta)"
      New-GPLink -Name $NombreGPO -Target $ruta -LinkEnabled Yes 

      Write-Output "La GPO $NombreGPO ha sido configurada y asignada a la OU $NombreOU"
    }
    else {
      Write-Output "La OU '$OU' ya existe en '$Dominio'. No se ha creado una nueva OU."
    }
    
  }
  elseif ($opcion -eq "no") {
    Write-Host "OK, continuamos..."
    
    #Asignamos la GPO a una unidad organizativa
    $NombreOU = Read-Host "Inserte el nombre de unidad organizativa donde se aplicará la GPO"
    $ruta = 'OU=' + $NombreOU + ',' + $Dominio
    $RutaGPC = "CN={$(Get-GPO -Name $NombreGPO).Id},CN=Policies,CN=System,$($ruta)"
    New-GPLink -Name $NombreGPO -Target $ruta -LinkEnabled Yes 
    Write-Host $ruta

    Write-Output "La GPO $NombreGPO ha sido configurada y asignada a la OU $NombreOU" -ForegroundColor Green
  }
  else {
    Write-Host "Opción incorrecta. Ingrese "si" o "no"" -ForegroundColor Red
  }
}


#DEFINICION DE FUNCION PARA LISTAR GPO A PARTIR DE UNA UNIDAD ORGANIZATIVA
function ObtenerGPO {
  $NombreUnidadOrganizativa = Read-Host "Ingrese el nombre de la Unidad Organizativa"
  try {
    $OU = Get-ADOrganizationalUnit -Filter { Name -eq $NombreUnidadOrganizativa }
    $GPOs = Get-GPO -All -Domain $OU.DistinguishedName
    $GPOs | Select-Object DisplayName, CreationTime, ModificationTime
    Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host
  }
  catch {
    Write-Error "No se pudo encontrar la unidad organizativa $NombreUnidadOrganizativa"
    Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host
  }
}


#DEFINICIÓN DE LA FUNCIÓN PARA LA CREACIÓN DE RECURSOS COMPARTIDOS
function NuevoRecursoCompartido {
  
  #Listamos todos los grupos para saber cual tomaremos de referencia
  Get-ADGroup -Filter *
  Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkYellow; Read-Host

  #Solicita al usuario los valores necesarios para el recurso compartido
  $NombreRecurso = Read-Host "Ingresa el nombre del recurso compartido"
  $Descripcion = Read-Host "Ingresa una descripción para el recurso compartido"
  $Ruta = Read-Host "Ingresa la ruta completa de la carpeta que deseas compartir"
  $GrupoRoot = Read-Host "Ingresa el nombre del grupo que tendrá permisos de acceso completo"
  $GrupoEditor = Read-Host "Ingresa el nombre del grupo que tendrá permisos de lectura y escritura"
  $GrupoLector = Read-Host "Ingresa el nombre del grupo que tendrá permisos de lectura"

  #Crea el objeto para el recurso compartido
  $DOMINIO = "FINETCOMPANY\"
  $D_GrupoE = $DOMINIO + $GrupoEditor
  $D_GrupoL = $DOMINIO + $GrupoLector
  Write-Host $D_GrupoE
  Write-Host $D_GrupoL
  $parametros = @{
    Name         = $NombreRecurso
    Path         = $Ruta
    Description  = $Descripcion
    FullAccess   = $GrupoRoot
    ChangeAccess = $D_GrupoE
    ReadAccess   = $D_GrupoL
  }
  New-SmbShare @parametros -EncryptData $True
  #Comprobamos que el recurso compartido se haya creado correctamente
  Write-Host "EL RECURSO COMPARTIDO SE HA CREADO CON EL SIGUIENTE NOMBRE Y PERMISOS DE GRUPOS:" -ForegroundColor Green
  Get-SmbShareAccess -Name $NombreRecurso | Select-Object -Property Name, AccountName, AccessRight
  #Comprobamos que el recurso compartido se haya creado correctamente
  # Get-SmbShare -Name $NombreRecurso
  Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host
 
}


#DEFINICIÓN DE FUNCIÓN PARA LISTAR INFORMACIÓN ACERCA DE LOS RECURSOS COMPARTIDOS CREADOS
function InformacionRecursoCompartido {

  #Obtener la lista de recursos compartidos del sistema
  $Recursos = Get-WmiObject Win32_Share
  
  #Recorrer cada recurso compartido
  foreach ($recurso in $Recursos) {
      
    #Obtiene la lista de los grupos que tienen permisos para el recurso compartido
    $ACL = Get-Acl $recurso.Path
    $grupos = $ACL.Access.IdentityReference.Value
      
    #Imprime la información del recurso compartido y sus permisos
    Write-Host "Recurso compartido: $($recurso.Name)"
    Write-Host "Directorio compartido: $($recurso.Path)"
    Write-Host "Grupos con permisos:"
    #Recorrer cada grupo para ir imprimiendolo por pantalla
    foreach ($grupo in $grupos) {
      Write-Host "- $grupo"
    }
    Write-Host ""
  }
  Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host
}



#DEFINICION DEL MENU DE SELECCION PARA EL USUARIO
$buqle = $true
while ($buqle) {
  write-host "╔════════════════════════════════════════════╗" -ForegroundColor White
  write-host "╬══MENÚ PARA LA GESTIÓN DE ACTIVE DIRECTORY══╬" -ForegroundColor DarkCyan
  Write-host "╚════════════════════════════════════════════╝" -ForegroundColor White
  write-host " 1." -ForegroundColor White -NoNewLine
  Write-host "Exportar usuarios y grupos a un archivo .csv" -ForegroundColor DarkGray
  write-host " 2." -ForegroundColor White -NoNewLine
  Write-host "Actualizar los atributos de los usuarios en funcion de la informacion de un archivo .csv" -ForegroundColor DarkGray
  write-host " 3." -ForegroundColor White -NoNewLine
  Write-host "Crear y administrar directivas de grupo (GPO)" -ForegroundColor DarkGray
  write-host " 4." -ForegroundColor White -NoNewLine
  Write-host "Listar GPO a partir de una OU" -ForegroundColor DarkGray
  write-host " 5." -ForegroundColor White -NoNewLine
  Write-host "Creación de recursos compartidos" -ForegroundColor DarkGray
  write-host " 6." -ForegroundColor White -NoNewLine
  Write-host "Listar información acerca de los recursos compartidos creados" -ForegroundColor DarkGray
  write-host " 7." -ForegroundColor White -NoNewLine
  write-host "Mostrar todos los usuarios" -ForegroundColor DarkGray
  write-host " 8." -ForegroundColor White -NoNewLine
  write-host "Crear un nuevo grupo" -ForegroundColor DarkGray
  write-host " X." -ForegroundColor Red -NoNewLine
  write-host "Salir" -ForegroundColor Gray
  $eleccion = read-host "Seleccione una opción" 
  switch ($eleccion) {
    1 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 1"
      Write-host "||EXPORTAR USUARIOS Y GRUPOS A UN ARCHIVO .CSV||" -ForegroundColor Cyan
      ExportaraCSV
      Clear-Host
    }
    2 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 2"
      Write-host "||ACTUALIZAR LOS ATRIBUTOS DE LOS USUARIOS EN FUNCION DE LA INFORMACIÓN DE UN ARCHIVO .CSV||" -ForegroundColor Cyan
      CambiarPropiedadesUsuariosCSV
    }
    3 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 3" 
      Write-host "||CREAR Y ADMINISTRAR DIRECTIVAS DE GRUPO (GPO)||" -ForegroundColor Cyan
      GestionarGPO
    }
    4 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 4" 
      Write-host "||LISTAR GPO A PARTIR DE UNA OU||" -ForegroundColor Cyan
      ObtenerGPO
    }
    5 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 5"
      Write-host "||CREACIÓN DE RECURSOS COMPARTIDOS||" -ForegroundColor Cyan
      NuevoRecursoCompartido
    }
    6 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 6"
      Write-host "||LISTAR INFORMACIÓN ACERCA DE LOS RECURSOS COMPARTIDOS CREADOS||" -ForegroundColor Cyan
      InformacionRecursoCompartido
    }
    7 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 7"
      Write-host "||MOSTRAR TODOS LOS USUARIOS||" -ForegroundColor Cyan
      Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host 
      Clear-Host
    }
    8 {
      Clear-host
      Write-host "Usted ha seleccionado la opción 8"
      Write-host "||CREAR UN NUEVO GRUPO||" -ForegroundColor Cyan
      Write-Host "Presione ↲ para continuar..." -ForegroundColor DarkRed; Read-Host 
      Clear-Host
    }
    ‘x’ { $buqle = $false }
    default { Write-Host "Eleccion invalida" -ForegroundColor Red }
  }
}