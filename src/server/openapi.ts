// Source for the in-app API screen. Keep summaries one line and examples accurate.

const error = (description: string) => ({
  description,
  content: { 'application/json': { schema: { $ref: '#/components/schemas/Error' } } },
})

const json = (schema: unknown) => ({ 'application/json': { schema } })
const ref = (name: string) => ({ $ref: `#/components/schemas/${name}` })
const idParam = { name: 'id', in: 'path', required: true, description: 'Schedule ID.', schema: { type: 'string' } }

const sampleMeta = {
  id: 'sch_3f9a1c',
  name: 'Flight Unit 2',
  createdAt: '2026-03-02T14:05:00.000Z',
  source: { kind: 'mspdi', filename: 'fu2.xml', importedAt: '2026-03-02T14:05:00.000Z' },
  taskCount: 142,
}

export const openApiDocument = {
  openapi: '3.1.0',
  info: {
    title: 'Streamline local API',
    version: '1.0.0',
    description:
      'Loopback-only API for importing, viewing and presenting schedules. Source data is read-only; only presentation and baselines can be changed. Errors are returned as { "error": { "message": "..." } }.',
  },
  servers: [{ url: 'http://127.0.0.1:4317' }],
  paths: {
    '/api/schedules': {
      get: {
        summary: 'List all imported schedules.',
        responses: { '200': { description: 'Schedule summaries.', content: json({ type: 'array', items: ref('ScheduleMeta') }) } },
        'x-example': { request: 'GET /api/schedules', response: [sampleMeta] },
      },
    },
    '/api/schedules/import/preview': {
      post: {
        summary: 'Parse an uploaded file and return a preview without saving anything.',
        description:
          'The body is the raw file. Send the file name in X-Filename (URL-encode non-ASCII). X-Mapping is an optional URL-encoded JSON column mapping for spreadsheets. X-Replace-Schedule-Id re-imports over an existing schedule and adds a reconciliation. Previews expire after 30 minutes.',
        parameters: [
          { name: 'X-Filename', in: 'header', required: true, schema: { type: 'string' } },
          { name: 'X-Mapping', in: 'header', required: false, schema: { type: 'string' } },
          { name: 'X-Replace-Schedule-Id', in: 'header', required: false, schema: { type: 'string' } },
        ],
        requestBody: { required: true, content: { 'application/octet-stream': { schema: { type: 'string', format: 'binary' } } } },
        responses: {
          '200': { description: 'The parsed preview.', content: json(ref('ImportPreview')) },
          '400': error('Missing file name, empty file or invalid mapping.'),
          '404': error('The schedule to replace was not found.'),
          '413': error('The file is larger than 25 MB.'),
          '422': error('The file cannot be imported.'),
        },
        'x-example': {
          request: 'POST /api/schedules/import/preview\nX-Filename: fu2.xml\n\n<file bytes>',
          response: { previewId: 'imp_8c1d', parsed: { kind: 'mspdi', filename: 'fu2.xml', name: 'Flight Unit 2', tasks: ['...'] }, summary: { rowCount: 142, laneCount: 6, dateRange: { start: '2026-01-05', finish: '2026-11-20' }, hasCriticality: true, hasFloat: true } },
        },
      },
    },
    '/api/schedules/import/accept': {
      post: {
        summary: 'Save a previewed import as a new schedule or as a replacement.',
        description: 'A replacement keeps the schedule ID, creation time and all presentation overrides, including ones whose task no longer exists.',
        requestBody: {
          required: true,
          content: json({ type: 'object', required: ['previewId'], properties: { previewId: { type: 'string' }, name: { type: 'string' } } }),
        },
        responses: {
          '201': { description: 'The new schedule.', content: json(ref('Schedule')) },
          '200': { description: 'The replaced schedule.', content: json(ref('Schedule')) },
          '400': error('The request is missing previewId or has an invalid name.'),
          '410': error('The preview has expired.'),
        },
        'x-example': { request: { previewId: 'imp_8c1d', name: 'Flight Unit 2' }, response: { id: 'sch_3f9a1c', name: 'Flight Unit 2', tasks: ['...'] } },
      },
    },
    '/api/schedules/{id}': {
      get: {
        summary: 'Get a schedule with its source tasks.',
        parameters: [idParam],
        responses: { '200': { description: 'The schedule.', content: json(ref('Schedule')) }, '404': error('Schedule not found.') },
        'x-example': { request: 'GET /api/schedules/sch_3f9a1c', response: { id: 'sch_3f9a1c', name: 'Flight Unit 2', tasks: ['...'] } },
      },
      patch: {
        summary: 'Rename a schedule. Only the schedule name label changes.',
        parameters: [idParam],
        requestBody: { required: true, content: json({ type: 'object', required: ['name'], properties: { name: { type: 'string' } } }) },
        responses: { '200': { description: 'Updated summary.', content: json(ref('ScheduleMeta')) }, '400': error('Invalid name.'), '404': error('Schedule not found.') },
        'x-example': { request: { name: 'FU2 baseline plan' }, response: { ...sampleMeta, name: 'FU2 baseline plan' } },
      },
      delete: {
        summary: 'Delete a schedule with its presentation and baselines.',
        parameters: [idParam],
        responses: { '204': { description: 'Deleted.' }, '404': error('Schedule not found.') },
        'x-example': { request: 'DELETE /api/schedules/sch_3f9a1c', response: '204 No Content' },
      },
    },
    '/api/schedules/{id}/view': {
      get: {
        summary: 'Resolve a view into ordered rows plus the timeline layout.',
        parameters: [
          idParam,
          { name: 'view', in: 'query', schema: { type: 'string', enum: ['executive', 'detailed', 'dependencies', 'workstreams', 'critical'], default: 'detailed' } },
          { name: 'page', in: 'query', schema: { type: 'integer', minimum: 1 } },
          { name: 'pageSize', in: 'query', schema: { type: 'integer', minimum: 1 } },
          { name: 'zoom', in: 'query', schema: { type: 'string', enum: ['day', 'week', 'month'], default: 'week' } },
          { name: 'collapsed', in: 'query', description: 'Comma-separated group keys.', schema: { type: 'string' } },
          { name: 'focusGroup', in: 'query', schema: { type: 'string' } },
          { name: 'text', in: 'query', description: 'Case-insensitive name filter.', schema: { type: 'string' } },
          { name: 'groups', in: 'query', description: 'Comma-separated group keys to include.', schema: { type: 'string' } },
          { name: 'criticality', in: 'query', description: 'Comma-separated: critical, not-critical, unavailable.', schema: { type: 'string' } },
          { name: 'overlapStart', in: 'query', description: 'YYYY-MM-DD. Requires overlapFinish.', schema: { type: 'string' } },
          { name: 'overlapFinish', in: 'query', description: 'YYYY-MM-DD. Requires overlapStart.', schema: { type: 'string' } },
          { name: 'compare', in: 'query', description: 'Baseline ID whose dates are included in the layout range.', schema: { type: 'string' } },
        ],
        responses: {
          '200': {
            description: 'Resolved rows and layout. layout.range is null when no visible task has a date.',
            content: json({ type: 'object', properties: { resolved: { type: 'object' }, layout: { type: 'object' } } }),
          },
          '400': error('A query value is invalid.'),
          '404': error('Schedule or baseline not found.'),
        },
        'x-example': {
          request: 'GET /api/schedules/sch_3f9a1c/view?view=executive&zoom=month',
          response: { resolved: { view: 'executive', rows: ['...'], page: 1, pageCount: 1 }, layout: { zoom: 'month', ppd: 4, range: { start: 20454, end: 20750 }, width: 1188, height: 116 } },
        },
      },
    },
    '/api/schedules/{id}/presentation': {
      get: {
        summary: 'Get presentation overrides for a schedule.',
        parameters: [idParam],
        responses: { '200': { description: 'The presentation, empty when none is saved.', content: json(ref('Presentation')) }, '404': error('Schedule not found.') },
        'x-example': { request: 'GET /api/schedules/sch_3f9a1c/presentation', response: { scheduleId: 'sch_3f9a1c', updatedAt: '1970-01-01T00:00:00.000Z', tasks: {} } },
      },
      patch: {
        summary: 'Merge presentation changes; source fields are rejected.',
        description:
          'Per task, fields merge into the existing override. A task value of null removes its override. A field set to null or "" is cleared. Allowed task fields: displayLabel, displayGroup, order, emphasis (none, highlight, muted). Source fields return 400.',
        parameters: [idParam],
        requestBody: {
          required: true,
          content: json({
            type: 'object',
            properties: {
              groupOrder: { type: ['array', 'null'], items: { type: 'string' } },
              theme: { type: ['string', 'null'] },
              tasks: { type: 'object', additionalProperties: { oneOf: [ref('TaskPresentation'), { type: 'null' }] } },
            },
          }),
        },
        responses: { '200': { description: 'The merged presentation.', content: json(ref('Presentation')) }, '400': error('Invalid field, source field, or unknown task.'), '404': error('Schedule not found.') },
        'x-example': { request: { tasks: { '12': { displayLabel: 'Engine hot fire', emphasis: 'highlight' } } }, response: { scheduleId: 'sch_3f9a1c', updatedAt: '2026-03-02T15:00:00.000Z', tasks: { '12': { displayLabel: 'Engine hot fire', emphasis: 'highlight' } } } },
      },
    },
    '/api/schedules/{id}/presentation/reset': {
      post: {
        summary: 'Clear presentation overrides for some or all tasks.',
        parameters: [idParam],
        requestBody: { required: false, content: json({ type: 'object', properties: { taskIds: { type: 'array', items: { type: 'string' } } } }) },
        responses: { '200': { description: 'The presentation after reset.', content: json(ref('Presentation')) }, '400': error('taskIds is not a list.'), '404': error('Schedule not found.') },
        'x-example': { request: { taskIds: ['12'] }, response: { scheduleId: 'sch_3f9a1c', updatedAt: '2026-03-02T15:05:00.000Z', tasks: {} } },
      },
    },
    '/api/schedules/{id}/baselines': {
      get: {
        summary: 'List saved baselines, newest first.',
        parameters: [idParam],
        responses: { '200': { description: 'Baselines.', content: json({ type: 'array', items: ref('Baseline') }) }, '404': error('Schedule not found.') },
        'x-example': { request: 'GET /api/schedules/sch_3f9a1c/baselines', response: [{ id: 'bl_77ab', scheduleId: 'sch_3f9a1c', name: 'Kickoff', createdAt: '2026-03-02T15:10:00.000Z', tasks: ['...'] }] },
      },
      post: {
        summary: 'Save the current source dates as a new immutable baseline.',
        parameters: [idParam],
        requestBody: { required: true, content: json({ type: 'object', required: ['name'], properties: { name: { type: 'string' } } }) },
        responses: { '201': { description: 'The new baseline.', content: json(ref('Baseline')) }, '400': error('Invalid name.'), '404': error('Schedule not found.') },
        'x-example': { request: { name: 'Kickoff' }, response: { id: 'bl_77ab', scheduleId: 'sch_3f9a1c', name: 'Kickoff', createdAt: '2026-03-02T15:10:00.000Z', tasks: ['...'] } },
      },
    },
    '/api/schedules/{id}/baselines/{baselineId}': {
      delete: {
        summary: 'Delete one baseline.',
        parameters: [idParam, { name: 'baselineId', in: 'path', required: true, schema: { type: 'string' } }],
        responses: { '204': { description: 'Deleted.' }, '404': error('Schedule or baseline not found.') },
        'x-example': { request: 'DELETE /api/schedules/sch_3f9a1c/baselines/bl_77ab', response: '204 No Content' },
      },
    },
    '/api/schedules/{id}/export': {
      post: {
        summary: 'Export source data as CSV or JSON (png, svg and pdf are not available yet).',
        description:
          'CSV columns: id,name,workstream,owner,start,finish,milestone,critical,float_days,parent_id,is_summary,predecessors. An empty CSV cell means the source did not provide the value. Predecessors are id:TYPE:lagDays joined by semicolons. JSON keeps nulls.',
        parameters: [idParam],
        requestBody: { required: true, content: json({ type: 'object', required: ['format'], properties: { format: { type: 'string', enum: ['png', 'svg', 'pdf', 'csv', 'json'] } } }) },
        responses: {
          '200': { description: 'The file, sent as an attachment.', content: { 'text/csv': { schema: { type: 'string' } }, 'application/json': { schema: { type: 'object' } } } },
          '400': error('Unknown format.'),
          '404': error('Schedule not found.'),
          '501': error('Not available yet.'),
        },
        'x-example': { request: { format: 'csv' }, response: 'id,name,workstream,owner,start,finish,milestone,critical,float_days,parent_id,is_summary,predecessors\r\n12,Engine hot fire,Propulsion,,2026-04-01,2026-04-03,false,critical,0,,false,10:FS:0' },
      },
    },
    '/api/openapi.json': {
      get: {
        summary: 'Return this OpenAPI document.',
        responses: { '200': { description: 'The OpenAPI 3.1 document.', content: json({ type: 'object' }) } },
        'x-example': { request: 'GET /api/openapi.json', response: { openapi: '3.1.0' } },
      },
    },
  },
  components: {
    schemas: {
      Error: { type: 'object', required: ['error'], properties: { error: { type: 'object', required: ['message'], properties: { message: { type: 'string' } } } } },
      ScheduleMeta: { type: 'object', required: ['id', 'name', 'createdAt', 'source', 'taskCount'], properties: { id: { type: 'string' }, name: { type: 'string' }, createdAt: { type: 'string' }, source: { type: 'object' }, taskCount: { type: 'integer' } } },
      Schedule: { type: 'object', required: ['id', 'name', 'createdAt', 'source', 'workstreams', 'tasks'], properties: { id: { type: 'string' }, name: { type: 'string' }, createdAt: { type: 'string' }, source: { type: 'object' }, workstreams: { type: 'array', items: { type: 'object' } }, tasks: { type: 'array', items: { type: 'object' } } } },
      TaskPresentation: { type: 'object', properties: { displayLabel: { type: 'string' }, displayGroup: { type: 'string' }, order: { type: 'number' }, emphasis: { type: 'string', enum: ['none', 'highlight', 'muted'] } } },
      Presentation: { type: 'object', required: ['scheduleId', 'updatedAt', 'tasks'], properties: { scheduleId: { type: 'string' }, updatedAt: { type: 'string' }, tasks: { type: 'object', additionalProperties: ref('TaskPresentation') }, groupOrder: { type: 'array', items: { type: 'string' } }, theme: { type: 'string' } } },
      Baseline: { type: 'object', required: ['id', 'scheduleId', 'name', 'createdAt', 'tasks'], properties: { id: { type: 'string' }, scheduleId: { type: 'string' }, name: { type: 'string' }, createdAt: { type: 'string' }, tasks: { type: 'array', items: { type: 'object' } } } },
      ImportPreview: {
        type: 'object',
        required: ['previewId', 'parsed', 'summary'],
        properties: { previewId: { type: 'string' }, parsed: { type: 'object' }, summary: { type: 'object' }, reconciliation: { type: 'object' } },
      },
    },
  },
}
