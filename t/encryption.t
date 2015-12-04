#!/usr/bin/env perl
use strict;
use warnings;
use Test::More;

use Spreadsheet::ParseXLSX;

my $parser = Spreadsheet::ParseXLSX->new(Password => '123q');
my $workbook = $parser->parse("t/data/encryption-agile-123q.xlsx");

my $worksheet;
my $cell;

$worksheet = $workbook->worksheet(0);
ok(defined($workbook));

$cell = $worksheet->get_cell(1, 1);
ok(defined($cell) && $cell->value() eq 'abcdefgABCDEFG');


open FH, "t/data/encryption-standard-default-password.xlsx";
$parser = Spreadsheet::ParseXLSX->new(Password => '');
$workbook = $parser->parse(\*FH);

ok(defined($workbook));

$worksheet = $workbook->worksheet(0);
$cell = $worksheet->get_cell(22, 8);
ok(defined($cell) && $cell->value() == 1911);

done_testing;
