#!/usr/bin/perl

use CGI qw(param);
my $name = param('name');
print "<center><h2>Useign forms and postiong data thew a perl script</h2><center>";
print "<hr>";
print "You just posted $name";