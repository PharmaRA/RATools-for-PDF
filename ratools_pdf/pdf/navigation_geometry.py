"""Update local PDF destinations without rebuilding bookmarks or link actions."""

import fitz


class _DestinationResize:
    def __init__(self, doc, transforms):
        self.doc = doc
        self.pdf = fitz._as_pdf_document(doc)
        self.mupdf = fitz.mupdf
        self.transforms = transforms
        self.pages = {doc.page_xref(index): index for index in range(doc.page_count)}
        self.patches = []
        self.visited = set()

    def get(self, obj, name):
        return self.mupdf.pdf_dict_gets(obj, name)

    def first_visit(self, obj, role):
        number = self.mupdf.pdf_to_num(obj)
        if number:
            key = role, number
            if key in self.visited:
                return False
            self.visited.add(key)
        return True

    def destination(self, value):
        m = self.mupdf
        if m.pdf_is_dict(value):
            array = self.get(value, 'D')
            if not m.pdf_is_array(array):
                return None
            replacement = self.destination(array)
            if replacement is None:
                return None
            copy = m.pdf_copy_dict(value)
            m.pdf_dict_put(copy, fitz.PDF_NAME('D'), replacement)
            return copy
        if not m.pdf_is_array(value) or m.pdf_array_len(value) < 2:
            return None  # Keep names; transform their definitions separately.
        page = m.pdf_array_get(value, 0)
        if m.pdf_is_indirect(page):
            page_index = self.pages.get(m.pdf_to_num(page), -1)
        elif m.pdf_is_int(page):
            page_index = m.pdf_to_int(page)
        else:
            return None
        if not 0 <= page_index < len(self.transforms):
            return None
        transform = self.transforms[page_index]
        if transform is None:
            return None
        mode = m.pdf_to_name(m.pdf_array_get(value, 1))
        axes = {'XYZ': {2: 'pdf_dx', 3: 'pdf_dy'},
                'FitH': {2: 'pdf_dy'}, 'FitBH': {2: 'pdf_dy'},
                'FitV': {2: 'pdf_dx'}, 'FitBV': {2: 'pdf_dx'},
                'FitR': {2: 'pdf_dx', 3: 'pdf_dy', 4: 'pdf_dx', 5: 'pdf_dy'}}.get(mode)
        if not axes:
            return None
        count = m.pdf_array_len(value)
        result = m.pdf_new_array(self.pdf, count)
        for index in range(count):
            item = m.pdf_array_get(value, index)
            if index in axes and m.pdf_is_number(item):
                m.pdf_array_push_real(result, m.pdf_to_real(item) * transform['scale'] + transform[axes[index]])
            else:
                m.pdf_array_push(result, item)  # Preserve null coordinates, zoom, and page reference.
        return result

    def queue_destination(self, owner, key):
        m = self.mupdf
        value = (m.pdf_array_get(owner, key) if isinstance(key, int) else
                 m.pdf_dict_gets(owner, key) if isinstance(key, str) else m.pdf_dict_get(owner, key))
        replacement = self.destination(value)
        if replacement is not None:
            self.patches.append((owner, key, replacement))

    def actions(self, initial):
        m = self.mupdf
        pending = [initial]
        while pending:
            action = pending.pop()
            if not self.first_visit(action, 'action'):
                continue
            if m.pdf_is_array(action):
                pending.extend(m.pdf_array_get(action, i) for i in range(m.pdf_array_len(action)))
            elif m.pdf_is_dict(action):
                if m.pdf_to_name(self.get(action, 'S')) == 'GoTo':
                    self.queue_destination(action, 'D')
                pending.append(self.get(action, 'Next'))

    def navigation(self, obj):
        self.queue_destination(obj, 'Dest')
        self.actions(self.get(obj, 'A'))
        additional = self.get(obj, 'AA')
        for index in range(self.mupdf.pdf_dict_len(additional)):
            self.actions(self.mupdf.pdf_dict_get_val(additional, index))

    def collect(self):
        m = self.mupdf
        catalog = m.pdf_dict_get(m.pdf_trailer(self.pdf), fitz.PDF_NAME('Root'))
        open_action = self.get(catalog, 'OpenAction')
        if m.pdf_is_array(open_action):
            self.queue_destination(catalog, 'OpenAction')
        else:
            self.actions(open_action)
        self.navigation(catalog)
        pending = [self.get(self.get(catalog, 'Outlines'), 'First')]
        while pending:
            outline = pending.pop()
            if not m.pdf_is_dict(outline) or not self.first_visit(outline, 'outline'):
                continue
            self.navigation(outline)
            pending.extend([self.get(outline, 'First'), self.get(outline, 'Next')])
        legacy = self.get(catalog, 'Dests')
        for index in range(m.pdf_dict_len(legacy)):
            self.queue_destination(legacy, m.pdf_dict_get_key(legacy, index))
        pending = [self.get(self.get(catalog, 'Names'), 'Dests')]
        while pending:
            tree = pending.pop()
            if not m.pdf_is_dict(tree) or not self.first_visit(tree, 'names'):
                continue
            names = self.get(tree, 'Names')
            for index in range(1, m.pdf_array_len(names), 2):
                self.queue_destination(names, index)
            children = self.get(tree, 'Kids')
            pending.extend(m.pdf_array_get(children, i) for i in range(m.pdf_array_len(children)))
        for page in self.doc:
            obj = fitz._as_pdf_page(page.this).obj()
            self.navigation(obj)
            annots = self.get(obj, 'Annots')
            for index in range(m.pdf_array_len(annots)):
                self.navigation(m.pdf_array_get(annots, index))

    def apply(self):
        # All replacements use original coordinates, including shared arrays.
        for owner, key, value in self.patches:
            if isinstance(key, int):
                self.mupdf.pdf_array_put(owner, key, value)
            elif isinstance(key, str):
                self.mupdf.pdf_dict_puts(owner, key, value)
            else:
                self.mupdf.pdf_dict_put(owner, key, value)


def resize_local_destinations(doc, transforms):
    operation = _DestinationResize(doc, transforms)
    operation.collect()
    operation.apply()
