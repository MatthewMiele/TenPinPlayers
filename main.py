import click
import player_analyser


@click.command()
@click.argument('filesdirectory', required=True, type=click.Path(exists=True, file_okay=False, resolve_path=True))
@click.argument('outputfilename', required=True, type=str)
@click.option('--leaguegapthreshold', type=int, default=5, help="How long should a gap between a players leagues be before highlighting")
def run(filesdirectory, outputfilename, leaguegapthreshold):
    """
    FILESDIRECTORY: Directory where the league files to be analysied are (xlsx files)\n
    OUTPUTFILENAME: File to save results to. Excluding extension.
    """
    player_analyser.analyse(filesdirectory, outputfilename, leaguegapthreshold)

if __name__ == '__main__':
    run()
