#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/page-Setup.xlsx');

my $ws1 = $wb->worksheet(0);

# Header/Footer Text
is($ws1->get_header, '&CHeader');
is($ws1->get_footer, '&CFooter');

# Margins
is($ws1->get_margin_header, '0.3');
is($ws1->get_margin_footer, '0.4');
is($ws1->get_margin_left,   '0.5');
is($ws1->get_margin_right,  '0.6');
is($ws1->get_margin_top,    '0.7');
is($ws1->get_margin_bottom, '0.8');

# Page Setup
is($ws1->get_print_scale, '75');
ok(!$ws1->is_portrait);
is($ws1->get_paper, '4');                               # Code for 11*17

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
