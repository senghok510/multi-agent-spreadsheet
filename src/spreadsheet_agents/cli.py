"""Command-line entry point for the project scaffold."""

import argparse


def main() -> None:
    parser = argparse.ArgumentParser(description="Multi-agent spreadsheet scaffold")
    parser.parse_args()
    parser.exit(message="Scaffold ready. Agent execution is not implemented yet.\n")


if __name__ == "__main__":
    main()
