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

    my ($version)    = $files->{workbook}->find_nodes('//fileVersion');
    my ($properties) = $files->{workbook}->find_nodes('//workbookPr');

    $workbook->{Version} = join('-',
        map { $version->att($_) } qw(appName lowestEdited)
    );
    $workbook->{Flag1904} = $properties->att('date1904') ? 1 : 0;

    $workbook->{FmtClass} = Spreadsheet::ParseExcel::FmtDefault->new; # XXX

    my $styles = $self->_parse_styles($files->{styles});

    $workbook->{Format}    = $styles->{Format};
    $workbook->{FormatStr} = $styles->{FormatStr};
    $workbook->{Font}      = $styles->{Font};

    $workbook->{PkgStr} = $self->_parse_shared_strings($files->{strings});

    # $workbook->{StandardWidth} = ...;

    # $workbook->{Author} = ...;

    # $workbook->{PrintArea} = ...;
    # $workbook->{PrintTitle} = ...;

    my @sheets = map {
        my $idx = $_->att('sheetId') - 1;
        my $sheet = Spreadsheet::ParseExcel::Worksheet->new(
            Name     => $_->att('name'),
            _Book    => $workbook,
            _SheetNo => $idx,
        );
        $self->_parse_sheet($sheet, $files->{sheets}[$idx]);
        $sheet
    } $files->{workbook}->find_nodes('//sheets/sheet');

    $workbook->{Worksheet}  = \@sheets;
    $workbook->{SheetCount} = scalar(@sheets);

    return $workbook;
}

sub _parse_sheet {
    my $self = shift;
    my ($sheet, $sheet_xml) = @_;

    # XXX need a fallback here, the dimension tag is optional
    my ($dimension) = $sheet_xml->find_nodes('//dimension');
    my ($topleft, $bottomright) = split ':', $dimension->att('ref');
    my ($rmin, $cmin) = $self->_cell_to_row_col($topleft);
    my ($rmax, $cmax) = $self->_cell_to_row_col($bottomright);

    $sheet->{MinRow} = $rmin;
    $sheet->{MinCol} = $cmin;
    $sheet->{MaxRow} = $rmax;
    $sheet->{MaxCol} = $cmax;

    for my $cell ($sheet_xml->find_nodes('//sheetData/row/c')) {
        my ($row, $col) = $self->_cell_to_row_col($cell->att('r'));
        my $val = $cell->first_child('v')->text;
        my $type = $cell->att('t') || 'n';

        my $long_type;
        if ($type eq 's') {
            $long_type = 'Text';
            $val = $sheet->{_Book}{PkgStr}[$val]{Text};
        }
        elsif ($type eq 'n') {
            $long_type = 'Numeric';
            $val = 0+$val;
        }
        elsif ($type eq 'd') {
            $long_type = 'Date';
        }
        else {
            die "unimplemented type $type"; # XXX
        }

        $sheet->{Cells}[$row][$col] = Spreadsheet::ParseExcel::Cell->new(
            Val     => $val,
            Type    => $long_type,
            # Format  => ...,
            ($cell->first_child('f')
                ? (Formula => $cell->first_child('f')->text)
                : ()),
        );
    }

    # ...
}

sub _parse_shared_strings {
    my $self = shift;
    my ($strings) = @_;

    return [
        map {
            { Text => $_->text } # XXX are Unicode, Rich, or Ext important?
        } $strings->find_nodes('//t')
    ];
}

sub _parse_styles {
    my $self = shift;
    my ($styles) = @_;

    my %format_str = map {
        $_->att('numFmtId') => $_->att('formatCode')
    } $styles->find_nodes('//numFmt');
    $format_str{0} = 'GENERAL'; # XXX others?

    my @font = map {
        Spreadsheet::ParseExcel::Font->new(
            Height         => 0+$_->first_child('sz')->att('val'),
            # Attr           => $iAttr,
            # Color          => $iCIdx,
            # Super          => $iSuper,
            # UnderlineStyle => $iUnderline,
            Name           => $_->first_child('name')->att('val'),

            # Bold      => $bBold,
            # Italic    => $bItalic,
            # Underline => $bUnderline,
            # Strikeout => $bStrikeout,
        )
    } $styles->find_nodes('//font');

    my @format = map {
        Spreadsheet::ParseExcel::Format->new(
            FontNo => 0+$_->att('fontId'),
            Font   => $font[$_->att('fontId')],
            FmtIdx => 0+$_->att('numFmtId'),

            # Lock     => $iLock,
            # Hidden   => $iHidden,
            # Style    => $iStyle,
            # Key123   => $i123,
            # AlignH   => $iAlH,
            # Wrap     => $iWrap,
            # AlignV   => $iAlV,
            # JustLast => $iJustL,
            # Rotate   => $iRotate,

            # Indent  => $iInd,
            # Shrink  => $iShrink,
            # Merge   => $iMerge,
            # ReadDir => $iReadDir,

            # BdrStyle => [ $iBdrSL, $iBdrSR,  $iBdrST, $iBdrSB ],
            # BdrColor => [ $iBdrCL, $iBdrCR,  $iBdrCT, $iBdrCB ],
            # BdrDiag  => [ $iBdrD,  $iBdrSD,  $iBdrCD ],
            # Fill     => [ $iFillP, $iFillCF, $iFillCB ],
        )
    } $styles->find_nodes('//cellXfs/xf');

    return {
        FormatStr => \%format_str,
        Font      => \@font,
        Format    => \@format,
    }
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

sub _cell_to_row_col {
    my $self = shift;
    my ($cell) = @_;

    my ($col, $row) = $cell =~ /([A-Z]+)([0-9]+)/;
    $col =~ tr/A-Z/0-9A-P/;
    $col = POSIX::strtol($col, 26);
    $row = $row - 1;

    return ($row, $col);
}

1;
