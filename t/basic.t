#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $filename = 't/data/Test.xlsx';

for my $file ($filename, do { open my $fh, '<', $filename or die; $fh }) {
    my $wb = Spreadsheet::ParseXLSX->new->parse($file);
    isa_ok($wb, 'Spreadsheet::ParseExcel::Workbook');

    is($wb->worksheet_count, 1);
    # is($workbook->get_filename, $filename); # XXX

    my $ws = $wb->worksheet(0);
    isa_ok($ws, 'Spreadsheet::ParseExcel::Worksheet');
    is($ws->get_name, 'Sheet1');
    is_deeply([$ws->row_range], [0, 5]);
    is_deeply([$ws->col_range], [0, 4]);

    {
        my $cell = $ws->get_cell(0, 0);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "Colored Cell");
        is($cell->value, "Colored Cell");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [1, '#9BBB59', '#FFFFFF']);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#FFFF00');
    }

    {
        my $cell = $ws->get_cell(0, 1);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "Wide Cell (25.00)");
        is($cell->value, "Wide Cell (25.00)");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    {
        my $cell = $ws->get_cell(0, 2);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "Bordered Cell w/ Text Wrap");
        is($cell->value, "Bordered Cell w/ Text Wrap");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok($format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(1) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    {
        my $cell = $ws->get_cell(0, 3);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "Middle Valigned");
        is($cell->value, "Middle Valigned");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 1);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    {
        my $cell = $ws->get_cell(0, 4);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "Right Aligned and text wrapped");
        is($cell->value, "Right Aligned and text wrapped");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 3);
        is($format->{AlignV}, 2);
        ok($format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    for my $i (0..4) {
        is($ws->get_cell(1, $i), undef);
    }

    {
        my $cell = $ws->get_cell(2, 0);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, 10);
        is($cell->value, 10);
        is($cell->type, 'Numeric');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    for my $i (1..4) {
        is($ws->get_cell(2, $i), undef);
    }

    {
        my $cell = $ws->get_cell(3, 0);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, 20);
        is($cell->value, 20);
        is($cell->type, 'Numeric');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    for my $i (1..2) {
        is($ws->get_cell(3, $i), undef);
    }

    {
        my $cell = $ws->get_cell(3, 3);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, 2.5);
        is($cell->value, "\$2.50");
        is($cell->type, 'Numeric');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    {
        my $cell = $ws->get_cell(3, 4);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "<< currency cell");
        is($cell->value, "<< currency cell");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 2);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#4BACC6');
    }

    {
        my $cell = $ws->get_cell(4, 0);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, 30);
        is($cell->value, 30);
        is($cell->type, 'Numeric');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    for my $i (1..4) {
        is($ws->get_cell(4, $i), undef);
    }

    {
        my $cell = $ws->get_cell(5, 0);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, 60);
        is($cell->value, 60);
        is($cell->type, 'Numeric');
        is($cell->{Formula}, 'SUM(A3:A5)');

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 0);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [0, undef, undef]);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#000000');
    }

    {
        my $cell = $ws->get_cell(5, 1);
        isa_ok($cell, 'Spreadsheet::ParseExcel::Cell');
        is($cell->unformatted, "<< formula cell");
        is($cell->value, "<< formula cell");
        is($cell->type, 'Text');
        is($cell->{Formula}, undef);

        my $format = $cell->get_format;
        isa_ok($format, 'Spreadsheet::ParseExcel::Format');
        is($format->{AlignH}, 3);
        is($format->{AlignV}, 2);
        ok(!$format->{Wrap});
        is_deeply($format->{Fill}, [1, '#EEECE1', '#FFFFFF']);
        is_deeply($format->{BdrStyle}, [(0) x 4]);
        is_deeply($format->{BdrColor}, [(undef) x 4]);
        is_deeply($format->{BdrDiag}, [0, 0, undef]);

        my $font = $format->{Font};
        isa_ok($font, 'Spreadsheet::ParseExcel::Font');
        is($font->{Name}, 'Calibri');
        is($font->{Height}, 12);
        is($font->{Color}, '#F79646');
    }

    for my $i (2..4) {
        is($ws->get_cell(5, $i), undef);
    }
}

done_testing;
