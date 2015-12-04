#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

{
    my $parser = Spreadsheet::ParseXLSX->new(Password => '123q');
    my $workbook = $parser->parse("t/data/encryption-agile-123q.xlsx");

    my $worksheet = $workbook->worksheet(0);
    my $cell = $worksheet->get_cell(1, 1);
    is($cell->value, 'abcdefgABCDEFG');
}

{
    open my $fh, "t/data/encryption-standard-default-password.xlsx" or die;
    my $parser = Spreadsheet::ParseXLSX->new(Password => '');
    my $workbook = $parser->parse($fh);

    my $worksheet = $workbook->worksheet(0);
    my $cell = $worksheet->get_cell(22, 8);
    is($cell->value, 1911);
}

done_testing;
