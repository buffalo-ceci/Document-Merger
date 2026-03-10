import React, { useState, useCallback } from 'react';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent,
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { FileText, File as FileIcon, GripVertical, Trash2, UploadCloud, Loader2 } from 'lucide-react';
import { mergeFiles } from './utils';

interface FileItemProps {
  id: string;
  file: File;
  onRemove: (id: string) => void;
}

const SortableFileItem: React.FC<FileItemProps> = ({ id, file, onRemove }) => {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 10 : 1,
  };

  const isPdf = file.name.toLowerCase().endsWith('.pdf');

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={`flex items-center gap-3 p-3 bg-white border rounded-lg shadow-sm mb-2 ${
        isDragging ? 'opacity-50 border-blue-500' : 'border-gray-200'
      }`}
    >
      <div
        {...attributes}
        {...listeners}
        className="cursor-grab p-1 text-gray-400 hover:text-gray-600"
      >
        <GripVertical size={20} />
      </div>
      <div className="flex-shrink-0">
        {isPdf ? (
          <FileIcon className="text-red-500" size={24} />
        ) : (
          <FileText className="text-blue-500" size={24} />
        )}
      </div>
      <div className="flex-grow min-w-0">
        <p className="text-sm font-medium text-gray-900 truncate">{file.name}</p>
        <p className="text-xs text-gray-500">{(file.size / 1024 / 1024).toFixed(2)} MB</p>
      </div>
      <button
        onClick={() => onRemove(id)}
        className="p-2 text-gray-400 hover:text-red-500 hover:bg-red-50 rounded-md transition-colors"
      >
        <Trash2 size={18} />
      </button>
    </div>
  );
}

export default function App() {
  const [files, setFiles] = useState<{ id: string; file: File }[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const sensors = useSensors(
    useSensor(PointerSensor),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    setError(null);

    const droppedFiles = Array.from(e.dataTransfer.files as Iterable<File> | ArrayLike<File>).filter(
      (f: File) => f.name.toLowerCase().endsWith('.pdf') || f.name.toLowerCase().endsWith('.docx')
    );

    if (droppedFiles.length === 0) {
      setError('Please drop only .pdf or .docx files.');
      return;
    }

    setFiles((prev) => [
      ...prev,
      ...droppedFiles.map((f: File) => ({ id: Math.random().toString(36).substring(7), file: f })),
    ]);
  }, []);

  const handleFileInput = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    if (e.target.files) {
      const selectedFiles = Array.from(e.target.files as Iterable<File> | ArrayLike<File>).filter(
        (f: File) => f.name.toLowerCase().endsWith('.pdf') || f.name.toLowerCase().endsWith('.docx')
      );
      
      if (selectedFiles.length === 0) {
        setError('Please select only .pdf or .docx files.');
        return;
      }

      setFiles((prev) => [
        ...prev,
        ...selectedFiles.map((f: File) => ({ id: Math.random().toString(36).substring(7), file: f })),
      ]);
    }
  }, []);

  const handleDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;

    if (over && active.id !== over.id) {
      setFiles((items) => {
        const oldIndex = items.findIndex((i) => i.id === active.id);
        const newIndex = items.findIndex((i) => i.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  const removeFile = (id: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== id));
  };

  const handleMerge = async () => {
    if (files.length === 0) return;
    
    setIsProcessing(true);
    setError(null);
    
    try {
      const mergedPdfBytes = await mergeFiles(files.map(f => f.file));
      
      const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'merged_document.pdf';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      setError('An error occurred while merging the files. Make sure the files are valid.');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-3xl mx-auto">
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-gray-900 tracking-tight">Document Merger</h1>
          <p className="mt-2 text-gray-600">
            Drag and drop PDF and DOCX files, reorder them, and merge into a single PDF.
          </p>
        </div>

        <div
          className={`relative border-2 border-dashed rounded-xl p-12 text-center transition-colors ${
            isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 bg-white hover:border-gray-400'
          }`}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <input
            type="file"
            multiple
            accept=".pdf,.docx"
            onChange={handleFileInput}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          />
          <UploadCloud className="mx-auto h-12 w-12 text-gray-400 mb-4" />
          <p className="text-lg font-medium text-gray-900">
            Click or drag files to this area to upload
          </p>
          <p className="text-sm text-gray-500 mt-1">Supports .pdf and .docx</p>
        </div>

        {error && (
          <div className="mt-4 p-4 bg-red-50 text-red-700 rounded-lg text-sm">
            {error}
          </div>
        )}

        {files.length > 0 && (
          <div className="mt-8">
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Files to Merge</h2>
            <DndContext
              sensors={sensors}
              collisionDetection={closestCenter}
              onDragEnd={handleDragEnd}
            >
              <SortableContext
                items={files.map((f) => f.id)}
                strategy={verticalListSortingStrategy}
              >
                <div className="space-y-2">
                  {files.map((f) => (
                    <SortableFileItem
                      key={f.id}
                      id={f.id}
                      file={f.file}
                      onRemove={removeFile}
                    />
                  ))}
                </div>
              </SortableContext>
            </DndContext>

            <div className="mt-8 flex justify-end">
              <button
                onClick={handleMerge}
                disabled={isProcessing}
                className={`flex items-center gap-2 px-6 py-3 rounded-lg text-white font-medium transition-colors ${
                  isProcessing
                    ? 'bg-blue-400 cursor-not-allowed'
                    : 'bg-blue-600 hover:bg-blue-700'
                }`}
              >
                {isProcessing ? (
                  <>
                    <Loader2 className="animate-spin" size={20} />
                    Processing...
                  </>
                ) : (
                  'Merge to PDF'
                )}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
