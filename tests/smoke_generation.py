from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app_core.services.registry_service import DOCUMENTS
from app_core.services.template_service import TemplateService
from app_core.services.generation_service import DocumentGenerationService

OUT = Path('test_generation_output')

def main():
    OUT.mkdir(exist_ok=True)
    ok = True
    for name, doc in DOCUMENTS.items():
        target = OUT / doc.engine_key
        target.mkdir(parents=True, exist_ok=True)
        print(f'TEST {name}: template={doc.default_template} exists={doc.default_template.exists()}')
        try:
            validated = TemplateService().validate(doc, doc.default_template)
            generated = DocumentGenerationService().generate(validated, doc.default_template, target)
            existing = [str(p) for p in generated if Path(p).exists()]
            print(' ->', existing)
            if len(existing) < 1:
                ok = False
        except Exception as exc:
            ok = False
            print(f'FAILED {name}: {type(exc).__name__}: {exc}')
    print('SMOKE_GENERATION_OK=', ok)
    raise SystemExit(0 if ok else 1)

if __name__ == '__main__':
    main()
