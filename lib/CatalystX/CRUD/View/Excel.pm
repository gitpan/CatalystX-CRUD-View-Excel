package CatalystX::CRUD::View::Excel;

use warnings;
use strict;
use base qw( Catalyst::View::Excel::Template::Plus CatalystX::CRUD );
use Data::Dump qw( dump );
use Path::Class;

our $VERSION = '0.01';

=head1 NAME

CatalystX::CRUD::View::Excel - view CRUD search/list results in Excel format

=head1 SYNOPSIS

=cut

=head2 new

Overrides base new() method to set default INCLUDE and TEMPLATE_EXTENSION
config values.

=cut

sub new {
    my ( $class, $c, $args ) = @_;
    my $self = $class->NEXT::new( $c, $args );
    $self->etp_config->{INCLUDE}        ||= [ $c->config->{root} ];
    $self->config->{TEMPLATE_EXTENSION} ||= 'tt';
    return $self;
}

=head2 process

Overrides base process() method to call get_filename()
and create template from results_template if a template file
does not exist.

=cut

sub process {
    my $self = shift;
    my $c    = shift;
    my @args = @_;

    my $template = $self->get_template_filename($c);

    ( defined $template )
        || die 'No template specified for rendering';

    # does $template exist? otherwise create one ad-hoc
    unless ( $self->template_exists($template) ) {
        $template = \( $self->results_template($c) );
    }

    my $etp_engine = $c->stash->{etp_engine} || $self->etp_engine;
    my $etp_config = $c->stash->{etp_config} || $self->etp_config;
    my $etp_params = $c->stash->{etp_params} || $self->etp_params;

    my $excel = $self->create_template_object(
        $c => (
            engine   => $etp_engine,
            template => $template,
            config   => $etp_config,
            params   => $etp_params,
        )
    );

    $excel->param( $self->get_template_params($c) );

    $c->response->content_type('application/x-msexcel');
    my $filename = $self->get_filename($c);
    $c->response->header( 'Content-Disposition',
        qq{attachment; filename="$filename"} );
    $c->response->body( $excel->output );
}

=head2 template_exists( I<path> )

Search the TT include path to see if I<path> really exists.

=cut

sub template_exists {
    my ( $self, $template ) = @_;
    for my $path ( @{ $self->etp_config->{INCLUDE} } ) {
        return 1 if -s file( $path, $template );
    }
    return 0;
}

=head2 get_template_filename( I<context> )

Overrides base method to change the default naming convention.
If C<template> is not set in stash(), then the default template
path is:

 $c->action . '.xls.' . $self->config->{TEMPLATE_EXTENSION}

C<TEMPLATE_EXTENSION> by default is C<tt>. You can alter that with the config()
method.

=cut

sub get_template_filename {
    my ( $self, $c ) = @_;
    $c->stash->{template}
        || ( $c->action . '.xls.' . $self->config->{TEMPLATE_EXTENSION} );
}

=head2 get_filename( I<context> )

Returns the name of the file to return in the response header
Content-Disposition.

=cut

sub get_filename {
    my ( $self, $c ) = @_;
    my $f = $c->action . '.xls';
    $f =~ s,/,_,g;
    return $f;
}

=head2 results_template( I<context> )

Returns results-specific template.

=cut

sub results_template {
    my ( $self, $c ) = @_;

    my $tmpl = <<TT;
<workbook>
    <worksheet name="[% c.controller.model_name.replace('\\W+','_') %]">
     [% myfields = c.controller.field_names %]
      <row>
     [% FOR fn = myfields %]
       <bold><cell>[% fn %]</cell></bold>
     [% END %]
      </row>
     [% FOR r = results.results %]
      <row>
         [% FOR fn = myfields %]
          <cell>[% r.\$fn %]</cell>
         [% END %]
      </row>
     [% END %]
    </worksheet>
</workbook>
TT

    return $tmpl;

}

=head1 AUTHOR

Peter Karman, C<< <karman at cpan dot org> >>

=head1 BUGS

Please report any bugs or feature requests to
C<bug-catalystx-crud-view-excel at rt.cpan.org>, or through the web interface at
L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=CatalystX-CRUD-View-Excel>.
I will be notified, and then you'll automatically be notified of progress on
your bug as I make changes.

=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc CatalystX::CRUD::View::Excel

You can also look for information at:

=over 4

=item * AnnoCPAN: Annotated CPAN documentation

L<http://annocpan.org/dist/CatalystX-CRUD-View-Excel>

=item * CPAN Ratings

L<http://cpanratings.perl.org/d/CatalystX-CRUD-View-Excel>

=item * RT: CPAN's request tracker

L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=CatalystX-CRUD-View-Excel>

=item * Search CPAN

L<http://search.cpan.org/dist/CatalystX-CRUD-View-Excel>

=back

=head1 ACKNOWLEDGEMENTS

The Minnesota Supercomputing Institute C<< http://www.msi.umn.edu/ >>
sponsored the development of this software.

=head1 COPYRIGHT & LICENSE

Copyright 2007 by the Regents of the University of Minnesota.

This program is free software; you can redistribute it and/or modify it
under the same terms as Perl itself.


=cut

1;
