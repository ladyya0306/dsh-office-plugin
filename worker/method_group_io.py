"""Indexed Office capabilities for a declared list of templates and outputs."""
from worker.method_io import OfficeCallbacks
from worker.office_worker import WorkerError
from pathlib import Path


class GroupCallbacks:
    def __init__(self, request):
        items = request['artifacts']
        if not isinstance(items, list) or not 1 <= len(items) <= 16:
            raise WorkerError('method_group_limit', '需1至16份明确目标')
        self.children = [OfficeCallbacks({**item, 'sources': request['sources']}) for item in items]
        self.items = items
        self.outputs = {}
        self.persist_generated = request.get('persist_generated', False)

    def check_conditions(self, _conditions):
        for child, item in zip(self.children, self.items):
            child.check_conditions(item['conditions'])

    def child(self, index):
        if isinstance(index, bool) or not isinstance(index, int) or not 0 <= index < len(self.children):
            raise WorkerError('method_group_target', '目标编号不在本批清单内')
        return self.children[index]

    def read_target(self, target_index=0):
        return self.child(target_index).read_target()

    def write_cells(self, changes, name, table_ranges=None, target_index=0):
        child = self.child(target_index)
        if name != self.items[target_index]['output_name'] or name in self.outputs:
            raise WorkerError('method_group_output', '产物名称需匹配本批目标编号且只生成一次')
        result = child.write_cells(changes, name, table_ranges)
        self.outputs[name] = child.outputs[name]
        if self.persist_generated:
            with Path(name).open('xb') as stream:
                stream.write(self.outputs[name])
        return result

    def functions(self):
        primary = self.children[0].functions()
        return {**primary, 'read_target': self.read_target, 'write_cells': self.write_cells}
