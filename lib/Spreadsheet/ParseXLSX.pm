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

    my $themes = $self->_parse_themes((values %{ $files->{themes} })[0]); # XXX

    $workbook->{Color} = $themes->{Color};

    my $styles = $self->_parse_styles($workbook, $files->{styles});

    $workbook->{Format}    = $styles->{Format};
    $workbook->{FormatStr} = $styles->{FormatStr};
    $workbook->{Font}      = $styles->{Font};

    $workbook->{PkgStr} = $self->_parse_shared_strings($files->{strings});

    # $workbook->{StandardWidth} = ...;

    # $workbook->{Author} = ...;

    # $workbook->{PrintArea} = ...;
    # $workbook->{PrintTitle} = ...;

    my @sheets = map {
        my $idx = $_->att('r:id');
        my $sheet = Spreadsheet::ParseExcel::Worksheet->new(
            Name     => $_->att('name'),
            _Book    => $workbook,
            _SheetNo => $idx,
        );
        $self->_parse_sheet($sheet, $files->{sheets}{$idx});
        $sheet
    } $files->{workbook}->find_nodes('//sheets/sheet');

    $workbook->{Worksheet}  = \@sheets;
    $workbook->{SheetCount} = scalar(@sheets);

    my ($node) = $files->{workbook}->find_nodes('//workbookView');
    my $selected = $node->att('activeTab');
    $workbook->{SelectedSheet} = defined($selected) ? 0+$selected : 0;

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
        my $val = $cell->first_child('v')
            ? $cell->first_child('v')->text
            : undef;
        my $type = $cell->att('t') || 'n';

        my $long_type;
        if ($type eq 's') {
            $long_type = 'Text';
            $val = $sheet->{_Book}{PkgStr}[$val]{Text};
        }
        elsif ($type eq 'n') {
            $long_type = 'Numeric';
            $val = defined($val) ? 0+$val : undef;
        }
        elsif ($type eq 'd') {
            $long_type = 'Date';
        }
        else {
            die "unimplemented type $type"; # XXX
        }

        $sheet->{Cells}[$row][$col] = Spreadsheet::ParseExcel::Cell->new(
            Val    => $val,
            Type   => $long_type,
            Format => $sheet->{_Book}{Format}[$cell->att('s') || 0],
            ($cell->first_child('f')
                ? (Formula => $cell->first_child('f')->text)
                : ()),
        );
    }

    my @column_widths;
    my @row_heights;

    my ($format) = $sheet_xml->find_nodes('//sheetFormatPr');
    my $default_row_height = $format->att('defaultRowHeight') || 15;
    my $default_column_width = $format->att('baseColWidth') || 10;

    for my $col ($sheet_xml->find_nodes('//col')) {
        $column_widths[$col->att('min') - 1] = $col->att('width');
    }

    for my $row ($sheet_xml->find_nodes('//row')) {
        $row_heights[$row->att('r') - 1] = $row->att('ht');
    }

    $sheet->{DefRowHeight} = 0+$default_row_height;
    $sheet->{DefColWidth} = 0+$default_column_width;
    $sheet->{RowHeight} = [
        map { defined $_ ? 0+$_ : 0+$default_row_height } @row_heights
    ];
    $sheet->{ColWidth} = [
        map { defined $_ ? 0+$_ : 0+$default_column_width } @column_widths
    ];

    my ($selection) = $sheet_xml->find_nodes('//selection');
    my $cell = $selection->att('activeCell');

    $sheet->{Selection} = [ $self->_cell_to_row_col($cell) ];
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

sub _parse_themes {
    my $self = shift;
    my ($themes) = @_;

    my @color = map {
        $_->name eq 'a:sysClr' ? $_->att('lastClr') : $_->att('val')
    } $themes->find_nodes('//a:clrScheme/*/*');

    # this shouldn't be necessary, but the documentation is wrong here
    # see http://stackoverflow.com/questions/2760976/theme-confusion-in-spreadsheetml
    ($color[0], $color[1]) = ($color[1], $color[0]);
    ($color[2], $color[3]) = ($color[3], $color[2]);

    return {
        Color => \@color,
    }
}

sub _parse_styles {
    my $self = shift;
    my ($workbook, $styles) = @_;

    my %halign = (
        center           => 2,
        centerContinuous => 6,
        distributed      => 7,
        fill             => 4,
        general          => 0,
        justify          => 5,
        left             => 1,
        right            => 3,
    );

    my %valign = (
        bottom      => 2,
        center      => 1,
        distributed => 4,
        justify     => 3,
        top         => 0,
    );

    my %border = (
        dashDot          => 9,
        dashDotDot       => 11,
        dashed           => 3,
        dotted           => 4,
        double           => 6,
        hair             => 7,
        medium           => 2,
        mediumDashDot    => 10,
        mediumDashDotDot => 12,
        mediumDashed     => 8,
        none             => 0,
        slantDashDot     => 13,
        thick            => 5,
        thin             => 1,
    );

    my %fill = (
        darkDown        => 7,
        darkGray        => 3,
        darkGrid        => 9,
        darkHorizontal  => 5,
        darkTrellis     => 10,
        darkUp          => 8,
        darkVertical    => 6,
        gray0625        => 18,
        gray125         => 17,
        lightDown       => 13,
        lightGray       => 4,
        lightGrid       => 15,
        lightHorizontal => 11,
        lightTrellis    => 16,
        lightUp         => 14,
        lightVertical   => 12,
        mediumGray      => 2,
        none            => 0,
        solid           => 1,
    );

    my @fills = map {
        [
            $fill{$_->att('patternType')},
            $self->_color($workbook->{Color}, $_->first_child('fgColor')),
            $self->_color($workbook->{Color}, $_->first_child('bgColor')),
        ]
    } $styles->find_nodes('//fills/fill/patternFill');

    my @borders = map {
        my $border = $_;
        # XXX specs say "begin" and "end" rather than "left" and "right",
        # but... that's not what seems to be in the file itself (sigh)
        {
            colors => [
                map {
                    $self->_color(
                        $workbook->{Color},
                        $border->first_child($_)->first_child('color')
                    )
                } qw(left right top bottom)
            ],
            styles => [
                map {
                    $border{$border->first_child($_)->att('style') || 'none'}
                } qw(left right top bottom)
            ],
            diagonal => [
                0, # XXX ->att('diagonalDown') and ->att('diagonalUp')
                0, # XXX ->att('style')
                $self->_color(
                    $workbook->{Color},
                    $border->first_child('diagonal')->first_child('color')
                ),
            ],
        }
    } $styles->find_nodes('//borders/border');

    my %format_str = map {
        $_->att('numFmtId') => $_->att('formatCode')
    } $styles->find_nodes('//numFmt');
    $format_str{0} = 'GENERAL'; # XXX others?

    my @font = map {
        Spreadsheet::ParseExcel::Font->new(
            Height         => 0+$_->first_child('sz')->att('val'),
            # Attr           => $iAttr,
            # XXX not sure if there's a better way to keep the indexing stuff
            # intact rather than just going straight to #xxxxxx
            Color          => $self->_color(
                $workbook->{Color},
                $_->first_child('color')
            ),
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
        my $alignment = $_->first_child('alignment');
        Spreadsheet::ParseExcel::Format->new(
            IgnoreFont         => !$_->att('applyFont'),
            IgnoreFill         => !$_->att('applyFill'),
            IgnoreBorder       => !$_->att('applyBorder'),
            IgnoreAlignment    => !$_->att('applyAlignment'),
            IgnoreNumberFormat => !$_->att('applyNumberFormat'),

            FontNo => 0+$_->att('fontId'),
            Font   => $font[$_->att('fontId')],
            FmtIdx => 0+$_->att('numFmtId'),

            # Lock     => $iLock,
            # Hidden   => $iHidden,
            # Style    => $iStyle,
            # Key123   => $i123,
            AlignH => $alignment
                ? $halign{$alignment->att('horizontal') || 'general'}
                : 0,
            Wrap => $alignment
                ? $alignment->att('wrapText')
                : 0,
            AlignV => $alignment
                ? $valign{$alignment->att('vertical') || 'bottom'}
                : 2,
            # JustLast => $iJustL,
            # Rotate   => $iRotate,

            # Indent  => $iInd,
            # Shrink  => $iShrink,
            # Merge   => $iMerge,
            # ReadDir => $iReadDir,

            BdrStyle => $borders[$_->att('borderId')]{styles},
            BdrColor => $borders[$_->att('borderId')]{colors},
            BdrDiag  => $borders[$_->att('borderId')]{diagonal},
            Fill     => $fills[$_->att('fillId')],
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

    my %worksheet_xml = map {
        $_->att('Id') => $self->_parse_xml($zip, $path_base . $_->att('Target'))
    } $wb_rels->find_nodes(qq<//Relationship[\@Type="$type_base/worksheet"]>);

    my %themes_xml = map {
        $_->att('Id') => $self->_parse_xml($zip, $path_base . $_->att('Target'))
    } $wb_rels->find_nodes(qq<//Relationship[\@Type="$type_base/theme"]>);

    return {
        workbook => $wb_xml,
        strings  => $strings_xml,
        styles   => $styles_xml,
        sheets   => \%worksheet_xml,
        themes   => \%themes_xml,
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

sub _color {
    my $self = shift;
    my ($colors, $color_node) = @_;

    my $color; # XXX
    if ($color_node) {
        $color = '#000000' # XXX
            if $color_node->att('auto');
        $color = '#' . Spreadsheet::ParseExcel->ColorIdxToRGB( # XXX
            $color_node->att('indexed')
        ) if defined $color_node->att('indexed');
        $color = '#' . substr($color_node->att('rgb'), 2, 6)
            if defined $color_node->att('rgb');
        $color = '#' . $colors->[$color_node->att('theme')]
            if defined $color_node->att('theme');
        # XXX tint?
    }

    return $color;
}

1;
