"""Laporan generator — turns raw marketplace exports (Transaksi/Saldo/Fee) into
standardized Laporan workbooks (Invoice/Jual/Remit/Bonus + Final).

Importable as a package: the orchestration entry point is ``laporan.main.main``
(CLI) / ``laporan.main.run`` (programmatic). Run standalone with
``python -m laporan`` from the repo root, or in-process via the bot's
``python main.py --laporan`` (which calls ``laporan.main.main`` directly).
"""
