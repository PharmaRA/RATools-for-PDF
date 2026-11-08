"""Inspect and remove active PDF entries while preserving unrelated structure."""

import fitz


_DYNAMIC_ACTIONS = {'JavaScript', 'Rendition', 'Movie', 'Sound', 'GoTo3DView', 'RichMediaExecute'}
_DYNAMIC_ANNOTATIONS = {'3D', 'RichMedia', 'Movie', 'Sound'}


class _CleanupPlan:
    def __init__(self, doc):
        self.doc = doc
        self.m = fitz.mupdf
        self.pdf = fitz._as_pdf_document(doc)
        self.patches = []
        self.count = 0
        self.visited = set()

    def get(self, obj, name):
        return self.m.pdf_dict_gets(obj, name)

    def array(self, items):
        array = self.m.pdf_new_array(self.pdf, len(items))
        for item in items:
            self.m.pdf_array_push(array, item)
        return array

    def chain(self, action, active=frozenset(), depth=0):
        """Return safe actions in order; splice out dynamic actions, retaining Next."""
        m = self.m
        if depth > 128:
            raise ValueError('动态动作链过深，无法安全清理，请人工复核')
        number = m.pdf_to_num(action)
        if number and number in active:
            return [], True  # Break cycles in the replacement, never follow them forever.
        active = active | {number} if number else active
        if m.pdf_is_array(action):
            result, changed = [], False
            for index in range(m.pdf_array_len(action)):
                safe, item_changed = self.chain(m.pdf_array_get(action,index),active,depth+1)
                result.extend(safe)
                changed |= item_changed
            return result, changed
        if not m.pdf_is_dict(action):
            return [], False
        following, changed = self.chain(self.get(action,'Next'), active, depth+1)
        if m.pdf_to_name(self.get(action,'S')) in _DYNAMIC_ACTIONS:
            self.count += 1
            return following, True
        if not changed:
            return [action], False
        copy = m.pdf_copy_dict(action)
        if following:
            m.pdf_dict_puts(copy,'Next',following[0] if len(following)==1 else self.array(following))
        else:
            m.pdf_dict_dels(copy,'Next')
        return [copy], True

    def root_action(self, actions):
        if not actions:
            return None
        if len(actions)==1:
            return actions[0]
        # A/OpenAction/AA require one action dictionary, while Next allows an array.
        first = self.m.pdf_copy_dict(actions[0])
        following = self.get(first,'Next')
        if self.m.pdf_is_array(following):
            tail = [self.m.pdf_array_get(following,i) for i in range(self.m.pdf_array_len(following))]
        else:
            tail = [following] if self.m.pdf_is_dict(following) else []
        self.m.pdf_dict_puts(first,'Next',self.array([*tail,*actions[1:]]))
        return first

    def action_entry(self, owner, key):
        action = self.get(owner,key)
        if not self.m.pdf_is_dict(action):
            return
        actions, changed = self.chain(action)
        if changed:
            self.patches.append((owner,key,self.root_action(actions)))

    def node(self, obj):
        if not self.visit(obj,'node'):
            return
        self.action_entry(obj,'A')
        additional = self.get(obj,'AA')
        if self.m.pdf_is_dict(additional):
            copy = self.m.pdf_copy_dict(additional)
            changed = False
            for index in range(self.m.pdf_dict_len(additional)):
                key = self.m.pdf_dict_get_key(additional,index)
                action = self.m.pdf_dict_get_val(additional,index)
                actions, item_changed = self.chain(action)
                if not item_changed:
                    continue
                changed = True
                replacement = self.root_action(actions)
                if replacement is None:
                    self.m.pdf_dict_del(copy,key)
                else:
                    self.m.pdf_dict_put(copy,key,replacement)
            if changed:
                self.patches.append((obj,'AA',copy if self.m.pdf_dict_len(copy) else None))

    def visit(self, obj, role):
        if not self.m.pdf_is_dict(obj):
            return False
        number = self.m.pdf_to_num(obj)
        if number:
            key = (role,number)
            if key in self.visited:
                return False
            self.visited.add(key)
        return True

    def collect(self):
        m = self.m
        catalog = self.get(m.pdf_trailer(self.pdf),'Root')
        names = self.get(catalog,'Names')
        if not m.pdf_is_null(self.get(names,'JavaScript')):
            copy = m.pdf_copy_dict(names)
            m.pdf_dict_dels(copy,'JavaScript')
            self.patches.append((catalog,'Names',copy if m.pdf_dict_len(copy) else None))
            self.count += 1
        self.action_entry(catalog,'OpenAction')
        self.node(catalog)
        pending = [self.get(self.get(catalog,'Outlines'),'First')]
        while pending:
            outline = pending.pop()
            if self.visit(outline,'outline'):
                self.node(outline)
                pending.extend([self.get(outline,'First'),self.get(outline,'Next')])
        form = self.get(catalog,'AcroForm')
        fields = self.get(form,'Fields')
        pending = [m.pdf_array_get(fields,i) for i in range(m.pdf_array_len(fields))]
        while pending:
            field = pending.pop()
            if self.visit(field,'field'):
                self.node(field)
                kids = self.get(field,'Kids')
                pending.extend(m.pdf_array_get(kids,i) for i in range(m.pdf_array_len(kids)))
        for page in self.doc:
            obj = fitz._as_pdf_page(page.this).obj()
            self.node(obj)
            annots = self.get(obj,'Annots')
            retained, removed = [], False
            for index in range(m.pdf_array_len(annots)):
                annot = m.pdf_array_get(annots,index)
                subtype = m.pdf_to_name(self.get(annot,'Subtype'))
                parent_type = m.pdf_to_name(self.get(self.get(annot,'Parent'),'Subtype'))
                if subtype in _DYNAMIC_ANNOTATIONS or (subtype=='Popup' and parent_type in _DYNAMIC_ANNOTATIONS):
                    removed = True
                    self.count += 1
                else:
                    self.node(annot)
                    retained.append(annot)
            if removed:
                self.patches.append((obj,'Annots',self.array(retained)))

    def apply(self):
        for owner,key,value in self.patches:
            if value is None:
                self.m.pdf_dict_dels(owner,key)
            else:
                self.m.pdf_dict_puts(owner,key,value)


def has_dynamic_content(doc):
    plan = _CleanupPlan(doc)
    plan.collect()
    return plan.count > 0


def remove_dynamic_content(doc):
    plan = _CleanupPlan(doc)
    plan.collect()
    if plan.count:
        plan.apply()
    return plan.count
