<script setup lang="ts">
import { computed, ref, watch } from 'vue'
import { grantCloudAccess, listCloudGrants, removeCloudGrant, setCloudGrantRole, setCloudLinkAccess, type CloudGrant, type CloudLinkLevel } from '@/cloud/access'
import { linkAccess } from '@/cloud/accessMap'
import type { CloudBinding } from '@/cloud/types'
import { useLocale } from '@/composables/useLocale'

const props = defineProps<{ modelValue: boolean; binding: CloudBinding | null }>()
const emit = defineEmits<{ 'update:modelValue': [value: boolean] }>()
const { t } = useLocale()

const grants = ref<CloudGrant[]>([])
const people = computed(() => grants.value.filter((grant) => !grant.link))
const linkLevel = computed(() => linkAccess(grants.value))
const email = ref('')
const role = ref<'reader' | 'writer'>('reader')
const busy = ref(false)
const errorText = ref('')

async function load() {
  if (!props.binding) return
  busy.value = true
  errorText.value = ''
  try {
    grants.value = await listCloudGrants(props.binding)
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('access.failed')
  } finally {
    busy.value = false
  }
}

watch(() => props.modelValue, (open) => {
  if (open) void load()
})

async function addGrant() {
  if (!props.binding) return
  busy.value = true
  errorText.value = ''
  try {
    await grantCloudAccess(props.binding, email.value, role.value)
    email.value = ''
    await load()
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('access.failed')
    busy.value = false
  }
}

async function changeRole(grant: CloudGrant, next: 'reader' | 'writer') {
  if (!props.binding || next === grant.role) return
  busy.value = true
  errorText.value = ''
  try {
    await setCloudGrantRole(props.binding, grant.id, next)
    await load()
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('access.failed')
    busy.value = false
  }
}

async function remove(grant: CloudGrant) {
  if (!props.binding) return
  busy.value = true
  errorText.value = ''
  try {
    await removeCloudGrant(props.binding, grant.id)
    await load()
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('access.failed')
    busy.value = false
  }
}

async function changeLink(next: CloudLinkLevel) {
  if (!props.binding || next === linkLevel.value) return
  busy.value = true
  errorText.value = ''
  try {
    await setCloudLinkAccess(props.binding, next)
    await load()
  } catch (error) {
    errorText.value = error instanceof Error ? error.message : t('access.failed')
    busy.value = false
  }
}

function onLink(event: Event) {
  const value = event.target instanceof HTMLSelectElement ? event.target.value : 'none'
  void changeLink(value === 'writer' || value === 'reader' ? value : 'none')
}

function onRole(grant: CloudGrant, event: Event) {
  const value = event.target instanceof HTMLSelectElement ? event.target.value : 'reader'
  void changeRole(grant, value === 'writer' ? 'writer' : 'reader')
}

function roleLabel(roleName: CloudGrant['role']) {
  if (roleName === 'owner') return t('access.owner')
  if (roleName === 'writer') return t('access.writer')
  return t('access.reader')
}
</script>

<template>
  <div v-if="modelValue" class="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" @click.self="emit('update:modelValue', false)">
    <div class="w-full max-w-md rounded-box border border-base-300 bg-base-100 p-4 shadow-xl">
      <div class="mb-3 flex items-center justify-between gap-2">
        <h2 class="truncate text-base font-semibold">{{ t('access.title') }}</h2>
        <button type="button" class="btn btn-ghost btn-xs" @click="emit('update:modelValue', false)">✕</button>
      </div>
      <p v-if="binding" class="mb-2 truncate text-xs opacity-60">{{ binding.file.name }}</p>
      <p v-if="errorText" class="mb-2 text-sm text-error">{{ errorText }}</p>
      <p v-if="busy" class="text-sm opacity-70">...</p>
      <label v-else class="mb-3 flex items-center gap-2 text-sm">
        <span class="min-w-0 flex-1">{{ t('access.link') }}</span>
        <select class="select select-bordered select-sm" :value="linkLevel" @change="onLink">
          <option value="none">{{ t('access.link.off') }}</option>
          <option value="reader">{{ t('access.reader') }}</option>
          <option value="writer">{{ t('access.full') }}</option>
        </select>
      </label>
      <ul v-if="!busy" class="mb-3 max-h-64 space-y-2 overflow-auto">
        <li v-for="grant in people" :key="grant.id" class="flex items-center gap-2 text-sm">
          <span class="min-w-0 flex-1 truncate">{{ grant.label }}</span>
          <select
            v-if="grant.editable"
            class="select select-bordered select-xs"
            :value="grant.role"
            @change="onRole(grant, $event)"
          >
            <option value="reader">{{ t('access.reader') }}</option>
            <option value="writer">{{ t('access.writer') }}</option>
          </select>
          <span v-else class="text-xs opacity-60">{{ roleLabel(grant.role) }}</span>
          <button v-if="grant.editable" type="button" class="btn btn-ghost btn-xs" @click="remove(grant)">{{ t('access.remove') }}</button>
        </li>
      </ul>
      <form class="flex gap-2" @submit.prevent="addGrant">
        <input v-model="email" type="email" required class="input input-bordered input-sm min-w-0 flex-1" :placeholder="t('access.email')" />
        <select v-model="role" class="select select-bordered select-sm">
          <option value="reader">{{ t('access.reader') }}</option>
          <option value="writer">{{ t('access.writer') }}</option>
        </select>
        <button type="submit" class="btn btn-primary btn-sm" :disabled="busy">{{ t('access.add') }}</button>
      </form>
    </div>
  </div>
</template>
