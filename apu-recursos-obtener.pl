#!/usr/bin/perl -w
# File     : apu-recursos-obtener.pl
# Author   : Patricio Rojas Ortiz
# version  : 2017-01-24 10:01:14
# revision : 2017-01-24 10:01:14
# Date     : 2017-01-23 10:33:58
#
# proposito: Este script perl, permite recorrer una planilla Microsoft excel
#            de analisis de precios unitarios (formato bien definido) y obtener
#            de ella el listado de recursos:
#            mano de obra, materiales e insumos, equipos y maquinarias,
#            herramientas y fungibles.
#
# License  : This perl script is under GNU General Public License.
#            More information you can found in
#            http://www.gnu.org/licenses/license-list.html#GNUGPL
#
# Plataforma : Los componentes requeridos son:
#            * Microsoft Windows system (Operating system)
#            * Microsoft office (tested with the version 2010)
#            * installed the DWIM Perl package (http://dwimperl.com/windows.html
#              descargar a instalar la version de
#               http://dwimperl.com/download/dwimperl-5.14.2.1-v7-32bit.exe
#              )
#
# ##############################################################################
use constant CONS_VERSION => '2017-01-25 13:58:55';

use strict;
use warnings;
use diagnostics;

use Cwd;                    # Used in sub_directory_default
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;

my ( $Excel_application, $Excel_book, $Excel_book_worksheets );
my ( $worksheets, $Sheet, $Sheet_total_nro_rows, $Sheet_row );
my ( $sl_filename_excel, $sl_tmp );

my ( $itemP, $item, $categoria, $unidad, $cantidad, $precio_unitario, $precio_total );
my ( @al_recursos_todos,
     @al_recursos_mano_obra,
     @al_recursos_equipos_y_herramientas,
     @al_recursos_materiales,
     @al_recursos_subcontratos,
     @al_recursos_otros_costos );

&sub_copyright();

if (not defined $ARGV[0] ){
  &sub_usage_help();
}

$sl_filename_excel = &sub_directory_default( $sl_filename_excel );
$sl_filename_excel = $sl_filename_excel. "\\".$ARGV[0];

&sub_verifyfile_exists( $sl_filename_excel ); # if file exists then continue

$Excel_application = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');
$Excel_application-> {Visible} = 1;       # 1 show Microsoft excel process
$Excel_application-> {DisplayAlerts} = 0; # This turns off the "This file already exists" message.

$Excel_book = $Excel_application->Workbooks->Open( $sl_filename_excel );
$Excel_book_worksheets = $Excel_book->Worksheets->count();

print "Leyendo archivo:";
foreach $worksheets (1..$Excel_book_worksheets){

	$Sheet = $Excel_book->Worksheets($worksheets);
  $Sheet->Activate();
	$Sheet_total_nro_rows= $Sheet->UsedRange->Rows->{'Count'};
	foreach $Sheet_row ( 1..$Sheet_total_nro_rows )
	{

# skip empty cells
    		next unless defined $Sheet->Cells($Sheet_row,2)->{'Value'};
        $itemP           = $Sheet->{'Name'} ;
        $item            = $Sheet->Cells($Sheet_row,  2)->{'Value'}; # Columna  2
        $categoria       = $Sheet->Cells($Sheet_row,  3)->{'Value'}; # Columna  3
        $unidad          = $Sheet->Cells($Sheet_row,  6)->{'Value'}; # Columna  6
        $cantidad        = $Sheet->Cells($Sheet_row,  7)->{'Value'}; # Columna  7
        $precio_unitario = $Sheet->Cells($Sheet_row,  9)->{'Value'}; # Columna  9
        $precio_total    = $Sheet->Cells($Sheet_row, 10)->{'Value'}; # Columna 10

# print  "$item $categoria $unidad $cantidad $precio_unitario $precio_total \n";

        if( defined $categoria and defined $unidad and defined $precio_unitario) {

          $sl_tmp          = $itemP."|". $item ."|". $categoria ."|". $unidad ."|". $cantidad ."|". $precio_unitario ."|". $precio_total;
print $sl_tmp."\n";
          push @al_recursos_todos, $sl_tmp;
          if( $item =~ /1./){	 push @al_recursos_mano_obra, $sl_tmp;              }
          if( $item =~ /2./){	 push @al_recursos_equipos_y_herramientas, $sl_tmp; }
          if( $item =~ /3./){	 push @al_recursos_materiales, $sl_tmp;             }
          if( $item =~ /4./){	 push @al_recursos_subcontratos, $sl_tmp;           }
          if( $item =~ /5./){	 push @al_recursos_otros_costos, $sl_tmp;           }

        }
	}
 }

# clean up after ourselves
$Excel_book->Close;

print "\nEscribiendo recursos-todos.csv\n";
$sl_filename_excel = ">recursos-todos.csv";
open ( OUTPUT_FILE, $sl_filename_excel ) or die "Error open file: $sl_filename_excel: $!";
print OUTPUT_FILE "Itemizado, Sub item, Descripcion, Unidad, Cantidad, Precio Unitario, Precio Total\n";
foreach  $sl_tmp ( @al_recursos_todos ){

  ($itemP, $item,$categoria,$unidad,$cantidad,$precio_unitario,$precio_total)=split(/\|/, $sl_tmp );
  if( defined $precio_unitario){
      print OUTPUT_FILE $itemP .",". $item.",".$categoria.",".$unidad.",".$cantidad.",".$precio_unitario.",".$precio_total. "\n";
  }

}
close OUTPUT_FILE;

print "Escribiendo recursos_mano_de_obra.csv\n";
$sl_filename_excel = ">recursos_mano_de_obra.csv";
open ( OUTPUT_FILE, $sl_filename_excel ) or die "Error open file: $sl_filename_excel: $!";
print OUTPUT_FILE "Itemizado, Sub item, Descripcion, Unidad, Cantidad, Precio Unitario, Precio Total\n";
foreach  $sl_tmp ( @al_recursos_mano_obra ){
  ($itemP, $item,$categoria,$unidad,$cantidad,$precio_unitario,$precio_total)=split(/\|/, $sl_tmp );
  if( defined $precio_unitario and $categoria){
      print OUTPUT_FILE $itemP .",". $item.",".$categoria.",".$unidad.",".$cantidad.",".$precio_unitario.",".$precio_total. "\n";
  }

}
close OUTPUT_FILE;

print "Escribiendo recursos_equipos-y-herramientas.csv\n";
$sl_filename_excel = ">recursos_equipos-y-herramientas.csv";
open ( OUTPUT_FILE, $sl_filename_excel ) or die "Error open file: $sl_filename_excel: $!";
print OUTPUT_FILE "Itemizado, Sub item, Descripcion, Unidad, Cantidad, Precio Unitario, Precio Total\n";
foreach  $sl_tmp ( @al_recursos_equipos_y_herramientas ){
  ($itemP, $item,$categoria,$unidad,$cantidad,$precio_unitario,$precio_total)=split(/\|/, $sl_tmp );
  if( defined $precio_unitario and $categoria){
      print OUTPUT_FILE $itemP .",". $item.",".$categoria.",".$unidad.",".$cantidad.",".$precio_unitario.",".$precio_total. "\n";
  }

}
close OUTPUT_FILE;

print "Escribiendo recursos_materiales.csv\n";
$sl_filename_excel = ">recursos_materiales.csv";
open ( OUTPUT_FILE, $sl_filename_excel ) or die "Error open file: $sl_filename_excel: $!";
print OUTPUT_FILE "Itemizado, Sub item, Descripcion, Unidad, Cantidad, Precio Unitario, Precio Total\n";
foreach  $sl_tmp ( @al_recursos_materiales ){
  ($itemP, $item,$categoria,$unidad,$cantidad,$precio_unitario,$precio_total)=split(/\|/, $sl_tmp );
  if( defined $precio_unitario and $categoria){
      print OUTPUT_FILE $itemP .",". $item.",".$categoria.",".$unidad.",".$cantidad.",".$precio_unitario.",".$precio_total. "\n";
  }

}
close OUTPUT_FILE;

print "Escribiendo recursos_subcontratos.csv\n";
$sl_filename_excel = ">recursos_subcontratos.csv";
open ( OUTPUT_FILE, $sl_filename_excel ) or die "Error open file: $sl_filename_excel: $!";
print OUTPUT_FILE "Itemizado, Sub item, Descripcion, Unidad, Cantidad, Precio Unitario, Precio Total\n";
foreach  $sl_tmp ( @al_recursos_subcontratos ){
  ($itemP, $item,$categoria,$unidad,$cantidad,$precio_unitario,$precio_total)=split(/\|/, $sl_tmp );
  if( defined $precio_unitario and $categoria){
      print OUTPUT_FILE $itemP .",". $item.",".$categoria.",".$unidad.",".$cantidad.",".$precio_unitario.",".$precio_total. "\n";
  }

}
close OUTPUT_FILE;

print "Escribiendo recursos_otros_costos.csv\n";
$sl_filename_excel = ">recursos_otros_costos.csv";
open ( OUTPUT_FILE, $sl_filename_excel ) or die "Error open file: $sl_filename_excel: $!";
print OUTPUT_FILE "Itemizado, Sub item, Descripcion, Unidad, Cantidad, Precio Unitario, Precio Total\n";
foreach  $sl_tmp ( @al_recursos_otros_costos ){
  ($itemP, $item,$categoria,$unidad,$cantidad,$precio_unitario,$precio_total)=split(/\|/, $sl_tmp );
  if( defined $precio_unitario and $categoria){
      print OUTPUT_FILE $itemP .",". $item.",".$categoria.",".$unidad.",".$cantidad.",".$precio_unitario.",".$precio_total. "\n";
  }

}
close OUTPUT_FILE;

print "Ok finalizando\n";
exit;
# ------------------------------------------------------------------------------
#                                                                    SUBROUTINES
# ------------------------------------------------------------------------------
sub sub_copyright(){

	print "\napu-recursos-obtener.pl V".CONS_VERSION.", Copyright 2017 appwebd.com\n\n";

}
# ------------------------------------------------------------------------------
sub sub_usage_help(){
  print "Uso: \n\napu-recursos-obtener.pl ANAPU_planilla_Excel\n\nANAPU_planilla_Excel es la Planilla de analisis de precios Unitarios en formato Microsoft Excel, ejemplo:\n\napu-recursos-obtener.pl APU.xlsx\n\nEste guion, tiene como salida de informacion cuatro archivos de recursos con la extension .csv.";
  exit;
}
# ------------------------------------------------------------------------------
sub sub_directory_default(){
  my $sl_dir_current = shift;

  if (	not defined $sl_dir_current or $sl_dir_current eq "" ){
     $sl_dir_current = getcwd; # getcwd (operate only in windows)

  }else{
    if($sl_dir_current eq ".." or $sl_dir_current eq "." ){
      $sl_dir_current = getcwd;
    }
  }
  return $sl_dir_current;

} #sub_directory_default
# -----------------------------------------------------------------------------
sub sub_verifyfile_exists(){
  my $sl_filename = shift;

  if (!(-e $sl_filename)) {
    print "Error, Archivo: $sl_filename no fue encontrado\n";
    exit;
  }
}
# ------------------------------------------------------------------------------
