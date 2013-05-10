package Spreadsheet::ParseXLSX;
use strict;
use warnings;

use Archive::Zip;
use Scalar::Util 'openhandle';
use Spreadsheet::ParseExcel;
use XML::Twig;

sub new {
    bless {}, shift;
}

sub parse {
    my $self = shift;
    my ($file) = @_;

    my $zip = Archive::Zip->new;
    if (openhandle($file)) {
        $zip->readFromFileHandle($file) == Archive::Zip::AZ_OK
            or die "Can't open filehandle as a zip file";
    }
    elsif (!ref($file)) {
        $zip->read($file) == Archive::Zip::AZ_OK
            or die "Can't open file '$file' as a zip file";
    }
    else {
        die "Argument to 'new' must be a filename or open filehandle";
    }

    return $self->_parse_workbook($zip);
}

sub _parse_workbook {
    my $self = shift;
    my ($zip) = @_;

    my $files = $self->_extract_files($zip);

    my $workbook = Spreadsheet::ParseExcel::Workbook->new;

    my @sheets = map {
        my $sheet = Spreadsheet::ParseExcel::Worksheet->new(
            Name     => $_->att('name'),
            _Book    => $workbook,
            _SheetNo => $_->att('sheetId') - 1,
        );
        $self->_parse_sheet($sheet, $files);
        $sheet
    } $files->{workbook}->find_nodes('//sheets/sheet');

    my ($version)    = $files->{workbook}->find_nodes('//fileVersion');
    my ($properties) = $files->{workbook}->find_nodes('//workbookPr');

    $workbook->{Version} = join('-',
        map { $version->att($_) } qw(appName lowestEdited)
    );
    $workbook->{Flag1904} = $properties->att('date1904') ? 1 : 0;

    $workbook->{FmtClass} = Spreadsheet::ParseExcel::FmtDefault->new; # XXX

    $workbook->{Worksheet}  = \@sheets;
    $workbook->{SheetCount} = scalar(@sheets);

    # $workbook->{Format}    = ...;
    # $workbook->{FormatStr} = ...;
    # $workbook->{Font}      = ...;

    # $workbook->{PkgStr} = ...;

    # $workbook->{StandardWidth} = ...;

    # $workbook->{Author} = ...;

    # $workbook->{PrintArea} = ...;
    # $workbook->{PrintTitle} = ...;

    return $workbook;
}

sub _parse_sheet {
    my $self = shift;
    my ($sheet, $files) = @_;

    # ...
}

sub _extract_files {
    my $self = shift;
    my ($zip) = @_;

    my $type_base =
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    my $rels = $self->_parse_xml(
        $zip,
        $self->_rels_for('')
    );
    my $wb_name = ($rels->find_nodes(
        qq<//Relationship[\@Type="$type_base/officeDocument"]>
    ))[0]->att('Target');
    my $wb_xml = $self->_parse_xml($zip, $wb_name);

    my $path_base = $self->_base_path_for($wb_name);
    my $wb_rels = $self->_parse_xml(
        $zip,
        $self->_rels_for($wb_name)
    );
    my $strings_xml = $self->_parse_xml(
        $zip,
        $path_base . ($wb_rels->find_nodes(
            qq<//Relationship[\@Type="$type_base/sharedStrings"]>
        ))[0]->att('Target')
    );
    my $styles_xml = $self->_parse_xml(
        $zip,
        $path_base . ($wb_rels->find_nodes(
            qq<//Relationship[\@Type="$type_base/styles"]>
        ))[0]->att('Target')
    );

    my @worksheet_xml = map {
        $self->_parse_xml($zip, $path_base . $_->att('Target'))
    } $wb_rels->find_nodes(qq<//Relationship[\@Type="$type_base/worksheet"]>);

    my @themes_xml = map {
        $self->_parse_xml($zip, $path_base . $_->att('Target'))
    } $wb_rels->find_nodes(qq<//Relationship[\@Type="$type_base/theme"]>);

    return {
        workbook => $wb_xml,
        strings  => $strings_xml,
        styles   => $styles_xml,
        sheets   => \@worksheet_xml,
        themes   => \@themes_xml,
    };
}

sub _parse_xml {
    my $self = shift;
    my ($zip, $subfile) = @_;

    my $member = $zip->memberNamed($subfile);
    die "no subfile named $subfile" unless $member;

    my $xml = XML::Twig->new;
    $xml->parse($member->contents);

    return $xml;
}

sub _rels_for {
    my $self = shift;
    my ($file) = @_;

    my @path = split '/', $file;
    my $name = pop @path;
    $name = '' unless defined $name;
    push @path, '_rels';
    push @path, "$name.rels";

    return join '/', @path;
}

sub _base_path_for {
    my $self = shift;
    my ($file) = @_;

    my @path = split '/', $file;
    pop @path;

    return join('/', @path) . '/';
}

1;
