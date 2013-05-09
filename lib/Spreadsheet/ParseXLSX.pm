package Spreadsheet::ParseXLSX;
use strict;
use warnings;

use Archive::Zip;
use Spreadsheet::ParseExcel;
use XML::Twig;

sub new {
    bless {}, shift;
}

sub parse {
    my $self = shift;
    my ($filename) = @_;

    $self->{Zip} = Archive::Zip->new;
    die "Can't open $filename as zip file"
        unless $self->{Zip}->read($filename) == Archive::Zip::AZ_OK;

    $self->{Workbook} = $self->_parse_workbook;
}

sub _parse_workbook {
    my $self = shift;

    my $files = $self->_extract_files;
    # ...
}

sub _extract_files {
    my $self = shift;

    my $type_base =
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    my $rels = $self->_parse_xml(
        $self->_rels_for('')
    );
    my $wb_name = ($rels->find_nodes(
        qq<//Relationship[\@Type="$type_base/officeDocument"]>
    ))[0]->att('Target');
    my $wb_xml = $self->_parse_xml($wb_name);

    my $path_base = $self->_base_path_for($wb_name);
    my $wb_rels = $self->_parse_xml(
        $self->_rels_for($wb_name)
    );
    my $strings_xml = $self->_parse_xml(
        $path_base . ($wb_rels->find_nodes(
            qq<//Relationship[\@Type="$type_base/sharedStrings"]>
        ))[0]->att('Target')
    );
    my $styles_xml = $self->_parse_xml(
        $path_base . ($wb_rels->find_nodes(
            qq<//Relationship[\@Type="$type_base/styles"]>
        ))[0]->att('Target')
    );

    my @worksheet_xml = map {
        $self->_parse_xml($path_base . $_->att('Target'))
    } $wb_rels->find_nodes(qq<//Relationship[\@Type="$type_base/worksheet"]>);

    my @themes_xml = map {
        $self->_parse_xml($path_base . $_->att('Target'))
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
    my ($subfile) = @_;

    my $member = $self->{Zip}->memberNamed($subfile);
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
