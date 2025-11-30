"""CLI interface for word-combiner utility."""
import click
from pathlib import Path
from typing import List

from word_combiner.combiner import combine_documents


@click.command()
@click.argument('input_files', nargs=-1, required=True, type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output',
    type=click.Path(path_type=Path),
    default=None,
    help='Output file path. If not specified, defaults to "combined.docx" in the current directory.'
)
@click.option(
    '--separator',
    default='page_break',
    type=click.Choice(['page_break', 'newline', 'none'], case_sensitive=False),
    help='Separator between documents: page_break (default), newline, or none.'
)
@click.option(
    '--sort',
    type=click.Choice(['name', 'date'], case_sensitive=False),
    default=None,
    help='Sort input files by name (alphabetical) or date (last modified). If not specified, files are processed in the order provided.'
)
def main(input_files: tuple, output: Path, separator: str, sort: str):
    """
    Combine multiple Word documents (.docx) into a single document.
    
    INPUT_FILES: One or more .docx files to combine
    """
    if not input_files:
        click.echo("Error: At least one input file is required.", err=True)
        raise click.Abort()
    
    # Validate all files are .docx
    for file_path in input_files:
        if not file_path.suffix.lower() == '.docx':
            click.echo(f"Error: {file_path} is not a .docx file.", err=True)
            raise click.Abort()
    
    # Set default output if not provided
    if output is None:
        output = Path.cwd() / 'combined.docx'
    
    # Ensure output has .docx extension
    if output.suffix.lower() != '.docx':
        output = output.with_suffix('.docx')
    
    # Sort files if requested
    files_list = list(input_files)
    if sort:
        if sort.lower() == 'name':
            files_list.sort(key=lambda p: p.name.lower())
        elif sort.lower() == 'date':
            files_list.sort(key=lambda p: p.stat().st_mtime)
    
    try:
        combine_documents(
            input_files=files_list,
            output_path=output,
            separator=separator
        )
        click.echo(f"Successfully combined {len(files_list)} document(s) into {output}")
    except Exception as e:
        click.echo(f"Error combining documents: {e}", err=True)
        raise click.Abort()


if __name__ == '__main__':
    main()
