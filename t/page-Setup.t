#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/page-Setup.xlsx');

my $ws1 = $wb->worksheet(0);

# Header/Footer Text
is($ws1->{header}, '&CHeader');
is($ws1->{footer}, '&CFooter');

# Margins
is($ws1->{pageMargins}{header}, '0.3');
is($ws1->{pageMargins}{footer}, '0.4');
is($ws1->{pageMargins}{left},   '0.5');
is($ws1->{pageMargins}{right},  '0.6');
is($ws1->{pageMargins}{top},    '0.7');
is($ws1->{pageMargins}{bottom}, '0.8');

# Page Setup
is($ws1->{pageSetup}{scale}, '75');
is($ws1->{pageSetup}{orientation}, 'landscape');
is($ws1->{pageSetup}{paperSize}, '4');                               # Code for 11*17

# Cell Border Formatting
is($ws1->get_cell(0,0)->get_format()->{'BdrDiag'}[0], 1);
is($ws1->get_cell(0,1)->get_format()->{'BdrDiag'}[0], 2);
is($ws1->get_cell(0,2)->get_format()->{'BdrDiag'}[0], 3);
is($ws1->get_cell(0,0)->get_format()->{'BdrDiag'}[1], 6);            # Double Line
is($ws1->get_cell(0,0)->get_format()->{'BdrDiag'}[2], '#FF0000');    # Red

is($ws1->get_cell(2,0)->get_format()->{'Rotate'}, 90);
is($ws1->get_cell(3,0)->get_format()->{'Shrink'}, 1);
is($ws1->get_cell(4,0)->get_format()->{'Indent'}, 1);

done_testing;
