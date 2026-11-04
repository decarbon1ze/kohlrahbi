"""
Command line interface for the context export (the explanatory text around the AHB tables).
"""

# Heavy submodules are imported lazily inside the command functions so that `--help` stays fast.

from pathlib import Path
from typing import Annotated

import typer
from rich.console import Console
from rich.panel import Panel

from kohlrahbi.cli_utils import bar_progress, prepare_command

console = Console()

context_app = typer.Typer(invoke_without_command=True)


@context_app.callback(invoke_without_command=True)
def context(
    edi_energy_mirror_path: Annotated[
        Path,
        typer.Option(
            "-eemp",
            "--edi-energy-mirror-path",
            help="The root path to the edi_energy_mirror repository.",
            exists=True,
            file_okay=False,
            dir_okay=True,
            resolve_path=True,
        ),
    ] = ...,  # type: ignore[assignment]
    output_path: Annotated[
        Path,
        typer.Option(
            "-o",
            "--output-path",
            help="Define the path where you want to save the generated files.",
            file_okay=False,
            dir_okay=True,
            resolve_path=True,
        ),
    ] = Path("output"),
    format_version: Annotated[
        str,
        typer.Option(
            "--format-version",
            help="Format version of the AHB documents, e.g. FV2310.",
        ),
    ] = ...,  # type: ignore[assignment]
    assume_yes: Annotated[
        bool,
        typer.Option(
            "-y",
            "--assume-yes",
            help="Confirm all prompts automatically.",
        ),
    ] = False,
    verbose: Annotated[
        bool,
        typer.Option(
            "-v",
            "--verbose",
            help="Enable verbose logging output.",
        ),
    ] = False,
) -> None:
    """
    Export the explanatory text around the AHB tables, with anchors to Prüfis, segments and conditions.
    """
    output_path, efv = prepare_command(
        console=console, verbose=verbose, output_path=output_path, assume_yes=assume_yes, format_version=format_version
    )

    from kohlrahbi.context import scrape_context

    with bar_progress(console) as progress:
        task = progress.add_task("Loading pruefi mapping...", total=None)

        def on_start(total: int) -> None:
            progress.update(task, description="Extracting context...", total=total)

        def on_file(name: str) -> None:
            progress.update(task, description=f"Extracted {name}")
            progress.advance(task)

        context_documents = scrape_context(
            basic_input_path=edi_energy_mirror_path,
            output_path=output_path,
            format_version=efv,
            on_start=on_start,
            on_file=on_file,
        )

    passages = sum(len(document.passages) for document in context_documents)
    console.print(
        Panel(
            f"[green]Documents:[/green] {len(context_documents)}\n"
            f"[green]Passages:[/green]  {passages}\n"
            f"[blue]Output:[/blue]    {output_path}",
            title="Context Export Complete",
            border_style="green",
        )
    )
