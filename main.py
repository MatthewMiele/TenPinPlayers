import click
import player_analiser


@click.command()
@click.argument('filesdirectory', required=True, type=click.Path(exists=True, file_okay=False, resolve_path=True))
def run(filesdirectory):
    """ FILESDIRECTORY: Directory where the league files to be analysied are (xlsx files)"""
    player_analiser.analise(filesdirectory)
    click.echo(filesdirectory)

if __name__ == '__main__':
    run()
