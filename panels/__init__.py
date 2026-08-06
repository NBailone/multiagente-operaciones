"""Panel mixins shared by the main ``App`` class.

Each mixin groups the UI/behaviour of one panel. Because they carry no
``__init__`` of their own, they blend into ``App`` via multiple inheritance
and rely on the shared attributes/methods living on the composed instance.
"""

from .impresion_mixin import ImpresionMixin
from .planillas_mixin import PlanillasMixin
from .correos_mixin import CorreosMixin
from .descarga_mixin import DescargaMixin
from .superauto_mixin import SuperAutoMixin
from .backup_mixin import BackupMixin

__all__ = [
    "ImpresionMixin", "PlanillasMixin", "CorreosMixin", "DescargaMixin",
    "SuperAutoMixin", "BackupMixin",
]