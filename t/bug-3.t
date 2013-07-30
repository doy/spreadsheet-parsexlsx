#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $wb = Spreadsheet::ParseXLSX->new->parse('t/data/bug-3.xlsx');
is($wb->worksheet_count, 1);

my $ws = $wb->worksheet(0);
is($ws->get_name, 'Sheet1');

is_deeply([$ws->row_range], [0, 1]);
is_deeply([$ws->col_range], [0, 2]);
is_deeply($ws->{Selection}, [1, 2]);

{
    my $cell = $ws->get_cell(0, 0);
    is($cell->value, "red");
    is($cell->type, 'Text');
    is($cell->get_format->{Font}{Color}, '#000000');
    is($cell->get_format->{Font}{Name}, 'Arial');
    is($cell->get_format->{Font}{Height}, '10');
}

{
    my $cell = $ws->get_cell(0, 1);
    is($cell->value, "blue");
    is($cell->type, 'Text');
    is($cell->get_format->{Font}{Color}, '#000000');
    is($cell->get_format->{Font}{Name}, 'Arial');
    is($cell->get_format->{Font}{Height}, '10');
}

{
    my $cell = $ws->get_cell(0, 2);
    is($cell->value, "green");
    is($cell->type, 'Text');
    is($cell->get_format->{Font}{Color}, '#000000');
    is($cell->get_format->{Font}{Name}, 'Arial');
    is($cell->get_format->{Font}{Height}, '10');
}

{
    my $cell = $ws->get_cell(1, 0);
    is($cell->value, "233");
    is($cell->type, 'Numeric');
    is($cell->get_format->{Font}{Color}, '#000000');
    is($cell->get_format->{Font}{Name}, 'Arial');
    is($cell->get_format->{Font}{Height}, '10');
}

{
    my $cell = $ws->get_cell(1, 1);
    is($cell->value, "444");
    is($cell->type, 'Numeric');
    is($cell->get_format->{Font}{Color}, '#000000');
    is($cell->get_format->{Font}{Name}, 'Arial');
    is($cell->get_format->{Font}{Height}, '10');
}

{
    my $cell = $ws->get_cell(1, 2);
    is($cell->value, "566");
    is($cell->type, 'Numeric');
    is($cell->get_format->{Font}{Color}, '#000000');
    is($cell->get_format->{Font}{Name}, 'Arial');
    is($cell->get_format->{Font}{Height}, '10');
}

done_testing;
